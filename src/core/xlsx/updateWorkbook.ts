/**
 * Atualização do modelo Excel.
 *
 * Em vez de regravar a planilha com uma biblioteca, editamos diretamente o XML dentro do
 * .xlsx: só a aba de dados (linhas 2 em diante, dimensão e posição de rolagem), a tabela
 * estruturada (intervalo) e, se necessário, UM estilo extra são alterados. Todo o resto do
 * pacote (tema, impressora, larguras, estilos, cabeçalhos...) permanece byte a byte igual.
 */

import { strFromU8, strToU8, zipSync, type Zippable } from "fflate";
import { COL, COLUMN_COUNT, FIRST_DATA_ROW } from "../constants";
import type { CalendarDate, PromoRecord } from "../models";
import { GenerationError, TemplateError } from "./errors";
import { parseSheetData, readCellValue, type CellValue } from "./sheetXml";
import { isDateStyle, parseStyles, quotePrefixVariant, serializeStyles } from "./styles";
import { analyzeTemplate, openPackage, text, type Files, type Layout } from "./template";
import { encodeXString, escapeXmlText, getAttr, numToCol, parseRange, setAttr } from "./xml";

const COLS = ["A", "B", "C", "D", "E", "F"] as const;

/** Número de série do Excel (sistema 1900, com o "bug" do ano bissexto de 1900). */
export function excelSerial(d: CalendarDate): number {
  const days = Math.round((Date.UTC(d.year, d.month - 1, d.day) - Date.UTC(1899, 11, 30)) / 86_400_000);
  return days < 61 ? days - 1 : days;
}

interface ExpectedCell {
  col: number;
  style: string | undefined;
  value: CellValue;
}

function styleAttr(s: string | undefined): string {
  return s === undefined ? "" : ` s="${s}"`;
}

/** Substitui o intervalo nas tags <sortState>/<sortCondition> que terminavam na última linha antiga. */
function shiftSortRefs(tableXml: string, oldEnd: number, newEnd: number): string {
  return tableXml.replace(/<(?:sortState|sortCondition)\b[^>]*>/g, (tag) => {
    const b = parseRange(getAttr(tag, "ref") ?? "");
    if (!b || b.maxRow !== oldEnd) return tag;
    const ref = `${numToCol(b.minCol)}${b.minRow}:${numToCol(b.maxCol)}${newEnd}`;
    return setAttr(tag, "ref", ref);
  });
}

export function updateWorkbook(template: Uint8Array, records: readonly PromoRecord[]): Uint8Array {
  if (records.length === 0) throw new GenerationError("Não há registros para gravar.");

  const files = openPackage(template);
  const layout = analyzeTemplate(files);
  const stylesXml = text(files, layout.stylesPath);
  if (!stylesXml) throw new TemplateError("O modelo não tem estilos (xl/styles.xml).");
  const styles = parseStyles(stylesXml);

  const { sheetData, sheetXml } = layout;
  const row1 = sheetData.rows.find((r) => r.r === 1);
  const row2 = sheetData.rows.find((r) => r.r === FIRST_DATA_ROW);
  if (!row1) throw new TemplateError("A linha de cabeçalhos do modelo não foi encontrada.");

  // Estilo-padrão: o da primeira linha de dados do modelo (linha 2).
  const templateStyle: Array<string | undefined> = COLS.map((_c, i) => {
    const cell = row2?.cells.find((c) => c.col === i + 1);
    return cell ? getAttr(cell.attrs, "s") : undefined;
  });
  for (const i of [COL.DATA_INICIAL, COL.DATA_FINAL]) {
    if (!isDateStyle(styles, Number(templateStyle[i] ?? 0))) {
      throw new TemplateError(
        "As colunas DATA INICIAL e DATA FINAL do modelo (D e E, linha 2) não estão formatadas como data.",
      );
    }
  }

  const rowAttrsFor = (n: number): string => {
    const base = row2?.attrs ?? ' r="2"';
    return /(^|\s)r="\d+"/.test(base) ? base.replace(/(^|\s)r="\d+"/, `$1r="${n}"`) : `${base} r="${n}"`;
  };

  // Células fora de A:F (ex.: observações em G) não pertencem aos dados e são preservadas.
  const extrasByRow = new Map<number, { attrs: string; cells: string[] }>();
  let maxCol: number = COLUMN_COUNT;
  let maxRow: number = FIRST_DATA_ROW + records.length - 1;
  for (const cell of row1.cells) maxCol = Math.max(maxCol, cell.col);
  for (const row of sheetData.rows) {
    if (row.r < FIRST_DATA_ROW) continue;
    const extra = row.cells.filter((c) => c.col > COLUMN_COUNT);
    if (extra.length === 0) continue;
    extrasByRow.set(row.r, { attrs: row.attrs, cells: extra.map((c) => c.raw) });
    maxCol = Math.max(maxCol, ...extra.map((c) => c.col));
    maxRow = Math.max(maxRow, row.r);
  }

  // --- novas linhas -------------------------------------------------------------
  const expected: ExpectedCell[][] = [];
  const rowsXml: string[] = [];
  const lastNewRow = FIRST_DATA_ROW + records.length - 1;

  const textCell = (col: string, row: number, s: string | undefined, value: string): { xml: string; style: string | undefined } => {
    let style = s;
    if (/^[=+\-@]/.test(value)) style = String(quotePrefixVariant(styles, Number(s ?? 0)));
    const xml =
      `<c r="${col}${row}"${styleAttr(style)} t="inlineStr"><is><t xml:space="preserve">` +
      `${escapeXmlText(encodeXString(value))}</t></is></c>`;
    return { xml, style };
  };
  const numCell = (col: string, row: number, s: string | undefined, n: number) =>
    `<c r="${col}${row}"${styleAttr(s)}><v>${n}</v></c>`;

  records.forEach((rec, i) => {
    const r = FIRST_DATA_ROW + i;
    const [sA, sB, sC, sD, sE, sF] = templateStyle;
    const nome = textCell("B", r, sB, rec.nome);
    const det = rec.detalhes === "" ? null : textCell("C", r, sC, rec.detalhes);
    const serialIni = excelSerial(rec.dataInicial);
    const serialFim = excelSerial(rec.dataFinal);
    const cells = [
      numCell("A", r, sA, rec.cod),
      nome.xml,
      det ? det.xml : `<c r="C${r}"${styleAttr(sC)}/>`,
      numCell("D", r, sD, serialIni),
      numCell("E", r, sE, serialFim),
      numCell("F", r, sF, rec.prioridade),
    ];
    const extras = extrasByRow.get(r)?.cells ?? [];
    rowsXml.push(`<row${rowAttrsFor(r)}>${cells.join("")}${extras.join("")}</row>`);
    expected.push([
      { col: 1, style: sA, value: rec.cod },
      { col: 2, style: nome.style, value: rec.nome },
      { col: 3, style: det ? det.style : sC, value: det ? rec.detalhes : null },
      { col: 4, style: sD, value: serialIni },
      { col: 5, style: sE, value: serialFim },
      { col: 6, style: sF, value: rec.prioridade },
    ]);
  });

  // Linhas abaixo dos novos dados que só têm conteúdo fora de A:F.
  for (const [r, { attrs, cells }] of [...extrasByRow].sort((a, b) => a[0] - b[0])) {
    if (r > lastNewRow) rowsXml.push(`<row${attrs}>${cells.join("")}</row>`);
  }

  const newSheetData = `<sheetData>${row1.raw}${rowsXml.join("")}</sheetData>`;
  const newRef = `A1:F${lastNewRow}`;

  // --- aba ------------------------------------------------------------------------
  let head = sheetXml.slice(0, sheetData.start);
  let tail = sheetXml.slice(sheetData.end);
  head = head.replace(/<dimension\b[^>]*\/>/, `<dimension ref="A1:${numToCol(maxCol)}${maxRow}"/>`);
  head = resetView(head);
  if (layout.sheetAutoFilter) tail = tail.replace(/<autoFilter\b[^>]*>/, (tag) => setAttr(tag, "ref", newRef));
  const newSheetXml = head + newSheetData + tail;

  // --- tabela estruturada -----------------------------------------------------------
  let newTableXml: string | undefined;
  if (layout.table) {
    const oldEnd = parseRange(layout.table.ref)!.maxRow;
    newTableXml = layout.table.xml
      .replace(/<table\b[^>]*>/, (tag) => setAttr(tag, "ref", newRef))
      .replace(/<autoFilter\b[^>]*>/, (tag) => setAttr(tag, "ref", newRef));
    newTableXml = shiftSortRefs(newTableXml, oldEnd, lastNewRow);
  }

  // --- pacote -----------------------------------------------------------------------
  const output: Zippable = {};
  for (const [name, data] of Object.entries(files)) output[name] = data;
  output[layout.sheetPath] = strToU8(newSheetXml);
  if (layout.table && newTableXml !== undefined) output[layout.table.path] = strToU8(newTableXml);
  if (styles.dirty) output[layout.stylesPath] = strToU8(serializeStyles(styles));

  let bytes: Uint8Array;
  try {
    bytes = zipSync(output, { level: 6 });
  } catch (exc) {
    throw new GenerationError(`Falha ao gerar o arquivo Excel (${exc instanceof Error ? exc.name : "erro"}).`);
  }

  verifyOutput(bytes, files, layout, expected, newRef, [
    layout.sheetPath,
    ...(layout.table ? [layout.table.path] : []),
    ...(styles.dirty ? [layout.stylesPath] : []),
  ]);
  return bytes;
}

/**
 * A janela do modelo abria rolada (A73) com seleção antiga (A2:F99); com outra quantidade de
 * linhas isso poderia abrir numa área vazia. Volta ao topo, a menos que haja painéis congelados.
 */
function resetView(head: string): string {
  if (/<pane\b/.test(head)) return head;
  return head
    .replace(/<sheetView\b[^>]*>/g, (tag) => tag.replace(/\stopLeftCell\s*=\s*(?:"[^"]*"|'[^']*')/, ""))
    .replace(/<selection\b[^>]*>/g, (tag) => setAttr(setAttr(tag, "activeCell", "A1"), "sqref", "A1"));
}

// --------------------------------------------------------------------------- //
// Conferência pós-geração
// --------------------------------------------------------------------------- //
function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function verifyOutput(
  output: Uint8Array,
  original: Files,
  layout: Layout,
  expected: ExpectedCell[][],
  expectedRef: string,
  modified: string[],
): void {
  const fail = (reason: string): never => {
    throw new GenerationError(`A conferência do arquivo gerado falhou: ${reason} O arquivo não foi disponibilizado.`);
  };

  let out: Files;
  try {
    out = openPackage(output);
  } catch {
    return fail("o arquivo gerado não pôde ser reaberto.");
  }

  const names = Object.keys(original);
  if (names.length !== Object.keys(out).length || names.some((n) => !out[n])) fail("a lista de partes do arquivo mudou.");
  for (const name of names) {
    if (!modified.includes(name) && !sameBytes(original[name], out[name])) fail(`a parte '${name}' foi alterada.`);
  }

  const sheetXml = strFromU8(out[layout.sheetPath]);
  const data = parseSheetData(sheetXml);
  if (!data) return fail("os dados da aba não puderam ser lidos.");
  const sst = layout.sst;

  const row1 = data.rows.find((r) => r.r === 1);
  const originalRow1 = layout.sheetData.rows.find((r) => r.r === 1);
  if (!row1 || !originalRow1 || row1.raw !== originalRow1.raw) fail("os cabeçalhos foram alterados.");

  const cols = (xml: string) => /<cols\b[\s\S]*?<\/cols>/.exec(xml)?.[0] ?? "";
  if (cols(sheetXml) !== cols(layout.sheetXml)) fail("as larguras das colunas foram alteradas.");

  const byRow = new Map(data.rows.map((r) => [r.r, r]));
  expected.forEach((cells, i) => {
    const r = FIRST_DATA_ROW + i;
    const row = byRow.get(r);
    if (!row) return fail(`a linha ${r} não foi encontrada.`);
    const found = row!.cells.filter((c) => c.col <= COLUMN_COUNT);
    if (found.length !== COLUMN_COUNT) fail(`a linha ${r} não tem as ${COLUMN_COUNT} colunas.`);
    cells.forEach((want, k) => {
      const cell = found[k];
      if (cell.col !== want.col || cell.row !== r) fail(`a linha ${r} tem células fora de ordem.`);
      if (getAttr(cell.attrs, "s") !== want.style) fail(`a formatação da linha ${r} difere do padrão do modelo.`);
      if (readCellValue(cell, sst) !== want.value) fail(`a linha ${r} não contém os valores esperados.`);
    });
  });
  for (const row of data.rows) {
    if (row.r > FIRST_DATA_ROW + expected.length - 1 && row.cells.some((c) => c.col <= COLUMN_COUNT)) {
      fail(`sobraram valores antigos na linha ${row.r}.`);
    }
  }

  if (layout.table) {
    const xml = strFromU8(out[layout.table.path]);
    const tag = /<table\b[^>]*>/.exec(xml)?.[0] ?? "";
    const af = /<autoFilter\b[^>]*>/.exec(xml)?.[0];
    if (getAttr(tag, "ref") !== expectedRef || (af && getAttr(af, "ref") !== expectedRef)) {
      fail("o intervalo da tabela estruturada está incorreto.");
    }
    if (getAttr(tag, "displayName") !== getAttr(/<table\b[^>]*>/.exec(layout.table.xml)![0], "displayName")) {
      fail("o nome da tabela estruturada foi alterado.");
    }
  }
}

/** Abertura do pacote .xlsx e análise (somente leitura) da estrutura do modelo. */

import { strFromU8, unzipSync } from "fflate";
import { COLUMN_COUNT, EXPECTED_HEADERS, FIRST_DATA_ROW } from "../constants";
import { normalizeHeaderName } from "../validators";
import { TemplateError } from "./errors";
import { parseSharedStrings, parseSheetData, readCellValue, type SheetData } from "./sheetXml";
import { getAttr, numToCol, parseRange, resolvePartPath } from "./xml";

export type Files = Record<string, Uint8Array>;

export interface TableInfo {
  path: string;
  xml: string;
  ref: string;
  name: string;
}

export interface Layout {
  files: Files;
  sheetName: string;
  sheetPath: string;
  sheetXml: string;
  sheetData: SheetData;
  sst: string[];
  /** Tabela estruturada que cobre A1:F..., se existir. */
  table: TableInfo | null;
  /** Autofiltro de planilha (sem tabela) cobrindo A1:F.... */
  sheetAutoFilter: boolean;
  stylesPath: string;
}

const NOT_XLSX =
  "O arquivo não é um .xlsx válido (pode estar corrompido, protegido por senha ou ser de outro formato, como .xls).";

export function openPackage(data: Uint8Array): Files {
  if (data.length === 0) throw new TemplateError("O arquivo enviado está vazio.");
  if (!(data.length > 3 && data[0] === 0x50 && data[1] === 0x4b)) throw new TemplateError(NOT_XLSX);
  let files: Files;
  try {
    files = unzipSync(data);
  } catch {
    throw new TemplateError(NOT_XLSX);
  }
  if (!files["xl/workbook.xml"] || !files["[Content_Types].xml"]) {
    throw new TemplateError("O arquivo não tem a estrutura de uma pasta de trabalho do Excel (.xlsx).");
  }
  if (files["xl/vbaProject.bin"]) {
    throw new TemplateError("Arquivos com macros (.xlsm) não são aceitos. Envie o modelo em .xlsx.");
  }
  return files;
}

export function text(files: Files, path: string): string | undefined {
  const data = files[path];
  return data ? strFromU8(data) : undefined;
}

function dirOf(path: string): string {
  const i = path.lastIndexOf("/");
  return i < 0 ? "" : path.slice(0, i);
}

function baseOf(path: string): string {
  return path.slice(path.lastIndexOf("/") + 1);
}

function relationships(files: Files, relsPath: string): Map<string, { target: string; type: string }> {
  const map = new Map<string, { target: string; type: string }>();
  const xml = text(files, relsPath);
  if (!xml) return map;
  for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = getAttr(m[0], "Id");
    const target = getAttr(m[0], "Target");
    if (id && target) map.set(id, { target, type: getAttr(m[0], "Type") ?? "" });
  }
  return map;
}

interface SheetRef {
  name: string;
  path: string;
}

function listWorksheets(files: Files): SheetRef[] {
  const workbook = text(files, "xl/workbook.xml")!;
  const rels = relationships(files, "xl/_rels/workbook.xml.rels");
  const sheets: SheetRef[] = [];
  for (const m of workbook.matchAll(/<sheet\b[^>]*>/g)) {
    const name = getAttr(m[0], "name");
    const rid = getAttr(m[0], "r:id");
    const rel = rid ? rels.get(rid) : undefined;
    if (!name || !rel || !rel.type.endsWith("/worksheet")) continue;
    const path = resolvePartPath("xl", rel.target);
    if (files[path]) sheets.push({ name, path });
  }
  return sheets;
}

function readHeaders(sheetData: SheetData, sst: string[]): string[] {
  const row = sheetData.rows.find((r) => r.r === 1);
  return Array.from({ length: COLUMN_COUNT }, (_v, i) => {
    const cell = row?.cells.find((c) => c.col === i + 1);
    const value = cell ? readCellValue(cell, sst) : null;
    return value === null ? "" : String(value);
  });
}

/** Valida a estrutura do modelo e localiza aba, tabela e estilos. Lança TemplateError. */
export function analyzeTemplate(files: Files): Layout {
  const workbookXml = text(files, "xl/workbook.xml")!;
  if (/<workbookPr\b[^>]*date1904\s*=\s*"(?:1|true)"/i.test(workbookXml)) {
    throw new TemplateError("O modelo usa o sistema de datas 1904 (Excel para Mac antigo), que não é suportado.");
  }

  const sst = parseSharedStrings(text(files, "xl/sharedStrings.xml"));
  const candidates: Array<{ ref: SheetRef; xml: string; data: SheetData }> = [];
  let firstHeaders: string[] | null = null;

  for (const ref of listWorksheets(files)) {
    const xml = text(files, ref.path)!;
    const data = parseSheetData(xml);
    if (!data) continue;
    const headers = readHeaders(data, sst);
    firstHeaders ??= headers;
    if (headers.every((h, i) => normalizeHeaderName(h) === EXPECTED_HEADERS[i])) candidates.push({ ref, xml, data });
  }

  if (candidates.length === 0) {
    throw new TemplateError(
      "Nenhuma aba do arquivo tem os cabeçalhos esperados na linha 1 " +
        `(${EXPECTED_HEADERS.join(", ")}). ` +
        `Cabeçalhos encontrados na primeira aba: ${firstHeaders?.filter(Boolean).join(", ") || "(vazio)"}.`,
    );
  }
  if (candidates.length > 1) {
    const names = candidates.map((c) => `'${c.ref.name}'`).join(", ");
    throw new TemplateError(
      `Mais de uma aba tem os cabeçalhos esperados (${names}); não é possível escolher com segurança.`,
    );
  }
  const { ref: sheet, xml: sheetXml, data: sheetData } = candidates[0];

  // --- tabelas estruturadas ---------------------------------------------------
  let table: TableInfo | null = null;
  const sheetRels = relationships(files, `${dirOf(sheet.path)}/_rels/${baseOf(sheet.path)}.rels`);
  for (const part of sheetXml.matchAll(/<tablePart\b[^>]*>/g)) {
    const rid = getAttr(part[0], "r:id");
    const rel = rid ? sheetRels.get(rid) : undefined;
    if (!rel) continue;
    const path = resolvePartPath(dirOf(sheet.path), rel.target);
    const xml = text(files, path);
    const tag = xml ? /<table\b[^>]*>/.exec(xml)?.[0] : undefined;
    if (!xml || !tag) continue;
    const tableRef = getAttr(tag, "ref") ?? "";
    const name = getAttr(tag, "displayName") ?? getAttr(tag, "name") ?? path;
    const b = parseRange(tableRef);
    if (!b) continue;
    if (b.minCol === 1 && b.minRow === 1 && b.maxCol === COLUMN_COUNT) {
      if (Number(getAttr(tag, "totalsRowCount") ?? 0) > 0) {
        throw new TemplateError(`A tabela '${name}' tem linha de totais, o que não é suportado.`);
      }
      const cols = [...xml.matchAll(/<tableColumn\b[^>]*>/g)].map((m) => normalizeHeaderName(getAttr(m[0], "name") ?? ""));
      if (cols.length !== COLUMN_COUNT || cols.some((c, i) => c !== EXPECTED_HEADERS[i])) {
        throw new TemplateError(`As colunas da tabela '${name}' não correspondem aos cabeçalhos esperados.`);
      }
      table = { path, xml, ref: tableRef, name };
    } else if (b.minCol <= COLUMN_COUNT && b.maxCol >= 1) {
      throw new TemplateError(
        `A aba tem outra tabela ('${name}', ${tableRef}) sobre as colunas A:F; não é possível atualizar com segurança.`,
      );
    }
  }

  // --- autofiltro de planilha (quando não há tabela) --------------------------
  let sheetAutoFilter = false;
  if (!table) {
    const tag = /<autoFilter\b[^>]*>/.exec(sheetXml.slice(sheetData.end))?.[0];
    const b = tag ? parseRange(getAttr(tag, "ref") ?? "") : null;
    if (b && b.minCol === 1 && b.minRow === 1 && b.maxCol === COLUMN_COUNT) sheetAutoFilter = true;
  }

  // --- verificações bloqueantes na área de dados ------------------------------
  for (const m of sheetXml.matchAll(/<mergeCell\b[^>]*>/g)) {
    const mergeRef = getAttr(m[0], "ref") ?? "";
    const b = parseRange(mergeRef);
    if (b && b.maxRow >= FIRST_DATA_ROW && b.minCol <= COLUMN_COUNT) {
      throw new TemplateError(
        `Há células mescladas na área de dados (${mergeRef}); não é possível atualizar com segurança.`,
      );
    }
  }
  for (const row of sheetData.rows) {
    if (row.r < FIRST_DATA_ROW) continue;
    for (const cell of row.cells) {
      if (cell.col <= COLUMN_COUNT && /<f[\s>/]/.test(cell.inner)) {
        throw new TemplateError(
          `A célula ${numToCol(cell.col)}${cell.row} contém fórmula na área de dados. ` +
            "A aplicação não substitui fórmulas; remova-a do modelo ou use outro arquivo.",
        );
      }
    }
  }

  return {
    files,
    sheetName: sheet.name,
    sheetPath: sheet.path,
    sheetXml,
    sheetData,
    sst,
    table,
    sheetAutoFilter,
    stylesPath: "xl/styles.xml",
  };
}

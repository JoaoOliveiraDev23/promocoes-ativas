/**
 * Constrói modelos .xlsx sintéticos com a MESMA estrutura do modelo real
 * (aba "Planilha1", Tabela1 estilo Medium16, estilos de célula, larguras, linhas vazias
 * formatadas sobrando no fim da tabela, janela rolada em A73, arquivo de impressora).
 * Os testes nunca usam dados reais.
 */
import { strToU8, zipSync, type Zippable } from "fflate";
import { escapeXmlText } from "../../src/core/xlsx/xml";

export interface ModelOptions {
  /** Quantidade de registros "antigos" preenchidos. */
  records?: number;
  /** Linhas de dados dentro da tabela (>= records); as demais ficam vazias e formatadas. */
  tableRows?: number;
  withTable?: boolean;
  sheetAutoFilter?: boolean;
  freeze?: boolean;
  sheetName?: string;
  headers?: string[];
  formulaInData?: boolean;
  mergeInData?: boolean;
  extraCell?: boolean; // célula G5 com observação (fora de A:F)
  extraEmptyStyledG32?: boolean;
  secondSheetWithSameHeaders?: boolean;
  otherTableOverlapping?: boolean;
  totalsRow?: boolean;
  dateStyle?: "date" | "general";
  date1904?: boolean;
  macro?: boolean;
  sortState?: boolean;
}

export const HEADERS = ["COD", "NOME", "DETALHES", "DATA INICIAL", "DATA FINAL", "PRIORIDADE"];
const NS = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"`;
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`;
const PRINTER_BYTES = new Uint8Array([0x01, 0x02, 0x03, 0xfe, 0xff, 0x00, 0x10, 0x20]);

export function sheetXmlFor(o: ModelOptions, shared: { idx: (s: string) => number }): string {
  const records = o.records ?? 5;
  const tableRows = Math.max(o.tableRows ?? 20, records);
  const dateS = o.dateStyle === "general" ? 3 : 4;
  const rows: string[] = [];
  const hdr = (o.headers ?? HEADERS)
    .map((h, i) => `<c r="${"ABCDEF"[i]}1" s="1" t="s"><v>${shared.idx(h)}</v></c>`)
    .join("");
  rows.push(`<row r="1" spans="1:6" x14ac:dyDescent="0.25">${hdr}</row>`);
  for (let i = 0; i < tableRows; i++) {
    const r = 2 + i;
    const filled = i < records;
    const cells = [
      filled ? `<c r="A${r}" s="3"><v>${9000 + i}</v></c>` : `<c r="A${r}" s="3"/>`,
      filled ? `<c r="B${r}" s="2" t="s"><v>${shared.idx(`Antiga ${i}`)}</v></c>` : `<c r="B${r}" s="2"/>`,
      filled ? `<c r="C${r}" s="2" t="s"><v>${shared.idx(`Detalhe antigo ${i}`)}</v></c>` : `<c r="C${r}" s="2"/>`,
      filled ? `<c r="D${r}" s="${dateS}"><v>${45658 + i}</v></c>` : `<c r="D${r}" s="${dateS}"/>`,
      filled ? `<c r="E${r}" s="${dateS}"><v>46387</v></c>` : `<c r="E${r}" s="${dateS}"/>`,
      filled ? `<c r="F${r}" s="3"><v>1</v></c>` : `<c r="F${r}" s="3"/>`,
    ];
    if (r === 2 && o.formulaInData) cells[0] = `<c r="A2" s="3"><f>1+1</f><v>2</v></c>`;
    if (r === 5 && o.extraCell) cells.push(`<c r="G5" t="inlineStr"><is><t>observação</t></is></c>`);
    if (r === 32 && o.extraEmptyStyledG32) cells.push(`<c r="G32" s="3"/>`);
    rows.push(`<row r="${r}" spans="1:6" x14ac:dyDescent="0.25">${cells.join("")}</row>`);
  }
  const endRow = 1 + tableRows;
  const view = o.freeze
    ? `<sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView>`
    : `<sheetView tabSelected="1" topLeftCell="A73" zoomScaleNormal="100" workbookViewId="0"><selection activeCell="A2" sqref="A2:F99"/></sheetView>`;
  const tail =
    (o.sheetAutoFilter ? `<autoFilter ref="A1:F${endRow}"/>` : "") +
    (o.mergeInData ? `<mergeCells count="1"><mergeCell ref="B3:C3"/></mergeCells>` : "") +
    `<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
    `<pageSetup paperSize="9" orientation="portrait" r:id="rId1"/>` +
    (o.withTable === false ? "" : `<tableParts count="${o.otherTableOverlapping ? 2 : 1}"><tablePart r:id="rId2"/>${o.otherTableOverlapping ? '<tablePart r:id="rId3"/>' : ""}</tableParts>`);
  return (
    DECL +
    `<worksheet ${NS}><dimension ref="A1:F${endRow}"/><sheetViews>${view}</sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>` +
    `<cols><col min="1" max="1" width="7.5703125" customWidth="1"/><col min="2" max="2" width="51.5703125" customWidth="1"/><col min="3" max="3" width="72.7109375" customWidth="1"/><col min="4" max="4" width="15.140625" customWidth="1"/><col min="5" max="5" width="13.5703125" customWidth="1"/><col min="6" max="6" width="15" customWidth="1"/></cols>` +
    `<sheetData>${rows.join("")}</sheetData>${tail}</worksheet>`
  );
}

function tableXml(o: ModelOptions, endRow: number, id: number, name: string, ref: string, cols: string[]): string {
  const totals = o.totalsRow && id === 1 ? ` totalsRowCount="1"` : ` totalsRowShown="0"`;
  return (
    DECL +
    `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${name}" displayName="${name}" ref="${ref}"${totals}>` +
    `<autoFilter ref="${ref}">${o.sortState && id === 1 ? `<sortState ref="A2:F${endRow}"><sortCondition ref="A2:A${endRow}"/></sortState>` : ""}</autoFilter>` +
    `<tableColumns count="${cols.length}">${cols.map((c, i) => `<tableColumn id="${i + 1}" name="${escapeXmlText(c)}"/>`).join("")}</tableColumns>` +
    `<tableStyleInfo name="TableStyleMedium16" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`
  );
}

/** Gera os bytes de um .xlsx sintético. */
export function buildModel(o: ModelOptions = {}): Uint8Array {
  const records = o.records ?? 5;
  const tableRows = Math.max(o.tableRows ?? 20, records);
  const endRow = 1 + tableRows;

  const strings: string[] = [];
  const shared = {
    idx(s: string): number {
      const i = strings.indexOf(s);
      if (i >= 0) return i;
      strings.push(s);
      return strings.length - 1;
    },
  };
  const sheet1 = sheetXmlFor(o, shared);
  const sst =
    DECL +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">` +
    strings.map((s) => `<si><t>${escapeXmlText(s)}</t></si>`).join("") +
    `</sst>`;

  const files: Zippable = {};
  files["[Content_Types].xml"] = strToU8(
    DECL +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`,
  );
  files["_rels/.rels"] = strToU8(
    DECL +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  const sheetName = o.sheetName ?? "Planilha1";
  files["xl/workbook.xml"] = strToU8(
    DECL +
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr${o.date1904 ? ' date1904="1"' : ""} defaultThemeVersion="164011"/><sheets><sheet name="${escapeXmlText(sheetName)}" sheetId="1" r:id="rId1"/>${o.secondSheetWithSameHeaders ? '<sheet name="Copia" sheetId="2" r:id="rId5"/>' : ""}</sheets></workbook>`,
  );
  files["xl/_rels/workbook.xml.rels"] = strToU8(
    DECL +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>${o.secondSheetWithSameHeaders ? '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' : ""}<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
  );
  files["xl/styles.xml"] = strToU8(
    DECL +
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Aptos Narrow"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="14" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf></cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
  );
  files["xl/sharedStrings.xml"] = strToU8(sst);
  files["xl/worksheets/sheet1.xml"] = strToU8(sheet1);
  if (o.secondSheetWithSameHeaders) {
    files["xl/worksheets/sheet2.xml"] = strToU8(sheetXmlFor({ ...o, withTable: false, sheetAutoFilter: false }, shared));
  }
  const rels = [
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>`,
  ];
  if (o.withTable !== false) {
    rels.push(`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>`);
    files["xl/tables/table1.xml"] = strToU8(tableXml(o, endRow, 1, "Tabela1", `A1:F${endRow}`, o.headers ?? HEADERS));
    if (o.otherTableOverlapping) {
      rels.push(`<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/>`);
      files["xl/tables/table2.xml"] = strToU8(tableXml(o, endRow, 2, "Outra", "B40:D45", ["X", "Y", "Z"]));
    }
  }
  files["xl/worksheets/_rels/sheet1.xml.rels"] = strToU8(
    DECL + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`,
  );
  files["xl/printerSettings/printerSettings1.bin"] = PRINTER_BYTES;
  files["xl/theme/theme1.xml"] = strToU8(DECL + `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Tema"/>`);
  if (o.macro) files["xl/vbaProject.bin"] = new Uint8Array([1, 2, 3]);
  return zipSync(files);
}

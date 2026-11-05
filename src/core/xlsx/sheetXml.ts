/** Leitura (somente texto/regex) do XML de uma aba do Excel. */

import { decodeXString, getAttr, unescapeXml } from "./xml";
import { colToNum } from "./xml";

export interface CellNode {
  col: number;
  row: number;
  /** Texto dos atributos da tag <c>, com espaço inicial (ex.: ` r="A1" s="4" t="s"`). */
  attrs: string;
  /** Conteúdo interno (vazio para células auto-fechadas). */
  inner: string;
  /** XML original completo da célula. */
  raw: string;
}

export interface RowNode {
  r: number;
  attrs: string;
  cells: CellNode[];
  raw: string;
}

export interface SheetData {
  /** Posição inicial/final de <sheetData>...</sheetData> dentro do XML da aba. */
  start: number;
  end: number;
  rows: RowNode[];
}

const SHEET_DATA_RE = /<sheetData\b[^>]*?(?:\/>|>([\s\S]*?)<\/sheetData>)/;
const ROW_RE = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
const CELL_RE = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;

export function parseSheetData(xml: string): SheetData | null {
  const m = SHEET_DATA_RE.exec(xml);
  if (!m) return null;
  const content = m[1] ?? "";
  const rows: RowNode[] = [];
  for (const rm of content.matchAll(ROW_RE)) {
    const attrs = rm[1];
    const r = Number(getAttr(attrs, "r") ?? NaN);
    if (!Number.isFinite(r)) continue;
    const cells: CellNode[] = [];
    for (const cm of (rm[2] ?? "").matchAll(CELL_RE)) {
      const ref = /^([A-Za-z]+)(\d+)$/.exec(getAttr(cm[1], "r") ?? "");
      if (!ref) continue;
      cells.push({ col: colToNum(ref[1]), row: Number(ref[2]), attrs: cm[1], inner: cm[2] ?? "", raw: cm[0] });
    }
    rows.push({ r, attrs, cells, raw: rm[0] });
  }
  return { start: m.index, end: m.index + m[0].length, rows };
}

/** Textos compartilhados (xl/sharedStrings.xml), já decodificados. */
export function parseSharedStrings(xml: string | undefined): string[] {
  if (!xml) return [];
  const out: string[] = [];
  for (const m of xml.matchAll(/<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g)) out.push(collectText(m[1] ?? ""));
  return out;
}

/** Junta o conteúdo de todos os <t> (ignorando fonética <rPh>). */
function collectText(xml: string): string {
  const clean = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, "");
  let text = "";
  for (const m of clean.matchAll(/<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/g)) text += unescapeXml(m[1] ?? "");
  return decodeXString(text);
}

export type CellValue = string | number | null;

/** Valor de uma célula (texto, número ou null se vazia). Datas voltam como número serial. */
export function readCellValue(cell: CellNode, sst: readonly string[]): CellValue {
  const t = getAttr(cell.attrs, "t");
  if (t === "inlineStr") {
    const is = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(cell.inner);
    return is ? collectText(is[1]) : null;
  }
  const v = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(cell.inner);
  if (!v) return null;
  const raw = unescapeXml(v[1]);
  if (t === "s") return sst[Number(raw)] ?? null;
  if (t === "str" || t === "e" || t === "b") return raw;
  const n = Number(raw);
  return Number.isNaN(n) ? raw : n;
}

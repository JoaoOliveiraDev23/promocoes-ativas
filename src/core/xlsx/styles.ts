/** Leitura mínima de xl/styles.xml: formatos de data e variante com quotePrefix. */

import { getAttr, setAttr } from "./xml";

const CELL_XFS_RE = /<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/;
const XF_RE = /<xf\b[^>]*?\/>|<xf\b[^>]*?>[\s\S]*?<\/xf>/g;
const NUM_FMT_RE = /<numFmt\b[^>]*?\/>/g;

/** Ids de formato numérico internos do Excel que representam data/hora. */
const BUILTIN_DATE_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55,
  56, 57, 58,
]);

export interface StylesInfo {
  xml: string;
  xfs: string[];
  numFmts: Map<number, string>;
  dirty: boolean;
  quoteVariants: Map<number, number>;
}

export function parseStyles(xml: string): StylesInfo {
  const numFmts = new Map<number, string>();
  for (const m of xml.matchAll(NUM_FMT_RE)) {
    const id = Number(getAttr(m[0], "numFmtId"));
    const code = getAttr(m[0], "formatCode");
    if (Number.isFinite(id) && code !== undefined) numFmts.set(id, code);
  }
  const section = CELL_XFS_RE.exec(xml);
  const xfs = section ? [...section[2].matchAll(XF_RE)].map((m) => m[0]) : [];
  return { xml, xfs, numFmts, dirty: false, quoteVariants: new Map() };
}

function isDateFormatCode(code: string): boolean {
  const stripped = code
    .replace(/"[^"]*"/g, "") // textos literais
    .replace(/\[[^\]]*\]/g, "") // cores/condições/locale
    .replace(/\\./g, ""); // caracteres escapados
  return /[dmyhs]/i.test(stripped) && !/^general$/i.test(stripped.trim());
}

/** True se o estilo (índice em cellXfs) usa um formato de data/hora. */
export function isDateStyle(info: StylesInfo, styleIndex: number): boolean {
  const xf = info.xfs[styleIndex];
  if (!xf) return false;
  const id = Number(getAttr(xf, "numFmtId") ?? 0);
  if (BUILTIN_DATE_IDS.has(id)) return true;
  const code = info.numFmts.get(id);
  return code !== undefined && isDateFormatCode(code);
}

/**
 * Devolve o índice de um estilo igual ao informado, porém com quotePrefix="1"
 * (a célula continua sendo tratada como texto mesmo se o usuário editá-la no Excel).
 */
export function quotePrefixVariant(info: StylesInfo, styleIndex: number): number {
  const cached = info.quoteVariants.get(styleIndex);
  if (cached !== undefined) return cached;
  const base = info.xfs[styleIndex];
  if (base === undefined) return styleIndex; // estilo desconhecido: não arriscar
  if (getAttr(base, "quotePrefix") === "1") return styleIndex;
  const open = /^<xf\b[^>]*?\/?>/.exec(base)![0];
  const variant = setAttr(open, "quotePrefix", "1") + base.slice(open.length);
  info.xfs.push(variant);
  info.dirty = true;
  const idx = info.xfs.length - 1;
  info.quoteVariants.set(styleIndex, idx);
  return idx;
}

/** XML de styles.xml com as variantes adicionadas (ou o original, se nada mudou). */
export function serializeStyles(info: StylesInfo): string {
  if (!info.dirty) return info.xml;
  const section = CELL_XFS_RE.exec(info.xml)!;
  const openTag = setAttr(`<cellXfs${section[1]}>`, "count", String(info.xfs.length));
  const replacement = `${openTag}${info.xfs.join("")}</cellXfs>`;
  return info.xml.slice(0, section.index) + replacement + info.xml.slice(section.index + section[0].length);
}

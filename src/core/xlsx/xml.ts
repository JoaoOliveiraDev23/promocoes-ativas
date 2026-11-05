/** Utilitários mínimos de XML/OOXML (sem dependências). Operam sobre texto, não sobre DOM. */

export function escapeXmlText(s: string): string {
  return s.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

export function escapeXmlAttr(s: string): string {
  return escapeXmlText(s).replaceAll('"', "&quot;");
}

export function unescapeXml(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (m, e: string) => {
    switch (e) {
      case "lt":
        return "<";
      case "gt":
        return ">";
      case "amp":
        return "&";
      case "quot":
        return '"';
      case "apos":
        return "'";
      default: {
        const code = e[1] === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        return Number.isFinite(code) && code <= 0x10ffff ? String.fromCodePoint(code) : m;
      }
    }
  });
}

/**
 * OOXML reserva a sequência `_xHHHH_` em textos (ex.: `_x000D_`). Para que um texto do usuário
 * que contenha essa sequência literal não seja reinterpretado pelo Excel, escapamos o "_" inicial.
 */
export function encodeXString(s: string): string {
  return s.replace(/_(x[0-9A-Fa-f]{4}_)/g, "_x005F_$1");
}

export function decodeXString(s: string): string {
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, hex: string) => String.fromCharCode(parseInt(hex, 16)));
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Lê o valor (já sem escapes XML) de um atributo dentro do texto de uma tag. */
export function getAttr(tag: string, name: string): string | undefined {
  const m = new RegExp(`(?:^|\\s)${escapeRegExp(name)}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag);
  if (!m) return undefined;
  return unescapeXml(m[1] ?? m[2] ?? "");
}

/** Define (ou cria) um atributo no texto de uma tag de abertura (`<tag ...>` ou `<tag .../>`). */
export function setAttr(tag: string, name: string, value: string): string {
  const re = new RegExp(`(\\s${escapeRegExp(name)}\\s*=\\s*)(?:"[^"]*"|'[^']*')`);
  const encoded = `"${escapeXmlAttr(value)}"`;
  if (re.test(tag)) return tag.replace(re, (_m, prefix: string) => `${prefix}${encoded}`);
  return tag.replace(/(\/?>)$/, ` ${name}=${encoded}$1`);
}

/** Remove um atributo do texto de uma tag. */
export function removeAttr(tag: string, name: string): string {
  return tag.replace(new RegExp(`\\s${escapeRegExp(name)}\\s*=\\s*(?:"[^"]*"|'[^']*')`), "");
}

// ----------------------------- colunas e intervalos ----------------------------- //
export function colToNum(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n: number): string {
  let s = "";
  let x = n;
  while (x > 0) {
    const rem = (x - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

export interface RangeBounds {
  minCol: number;
  minRow: number;
  maxCol: number;
  maxRow: number;
}

/** "A1:F125" (ou "B7") -> limites numéricos. Retorna null se inválido. */
export function parseRange(ref: string): RangeBounds | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(ref.trim());
  if (!m) return null;
  const minCol = colToNum(m[1]);
  const minRow = Number(m[2]);
  const maxCol = m[3] ? colToNum(m[3]) : minCol;
  const maxRow = m[4] ? Number(m[4]) : minRow;
  return { minCol, minRow, maxCol, maxRow };
}

/** Resolve um alvo de relacionamento ("../tables/table1.xml") relativo ao diretório base. */
export function resolvePartPath(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = baseDir.split("/").filter(Boolean);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== "." && seg !== "") parts.push(seg);
  }
  return parts.join("/");
}

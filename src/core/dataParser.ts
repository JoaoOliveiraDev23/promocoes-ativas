/** Leitura do texto colado do SQL Server (colunas separadas por tabulação). */

import { EXPECTED_HEADERS } from "./constants";
import type { ParseResult, PromoRecord, RawRow } from "./models";
import { crossCheck, looksLikeHeader, validateRow } from "./validators";

interface QuotedField {
  value: string;
  nextPos: number;
  newlines: number;
}

/**
 * Tenta ler um campo entre aspas (como o Excel faz ao colar texto).
 *
 * Só é considerado "entre aspas" se a aspa de fechamento for seguida de tabulação,
 * quebra de linha ou fim do texto. Caso contrário o campo é lido literalmente
 * (uma aspa solta no meio de um texto não altera nada).
 */
function readQuoted(text: string, pos: number): QuotedField | null {
  const n = text.length;
  let i = pos + 1;
  let out = "";
  let newlines = 0;
  while (i < n) {
    const ch = text[i];
    if (ch === '"') {
      if (i + 1 < n && text[i + 1] === '"') {
        out += '"'; // aspa escapada ("")
        i += 2;
        continue;
      }
      const end = i + 1;
      if (end >= n || "\t\r\n".includes(text[end])) {
        return { value: out.replaceAll("\r\n", "\n").replaceAll("\r", "\n"), nextPos: end, newlines };
      }
      return null;
    }
    if (ch === "\n") {
      newlines++;
    } else if (ch === "\r" && !(i + 1 < n && text[i + 1] === "\n")) {
      newlines++; // \r isolado; em \r\n a contagem é feita no \n
    }
    out += ch;
    i++;
  }
  return null;
}

/**
 * Divide o texto em linhas e campos (separador: tabulação).
 *  - Aceita quebras \n, \r\n e \r.
 *  - Não remove espaços dos campos.
 *  - Não cria linha extra quando o texto termina com quebra de linha.
 *  - `line` é o número da primeira linha física do registro (base 1).
 */
export function splitTsv(input: string): RawRow[] {
  const text = input.startsWith("\ufeff") ? input.slice(1) : input;
  const rows: RawRow[] = [];
  const n = text.length;
  let pos = 0;
  let line = 1;
  let rowLine = 1;
  let fields: string[] = [];

  for (;;) {
    let value: string | null = null;
    if (pos < n && text[pos] === '"') {
      const quoted = readQuoted(text, pos);
      if (quoted) {
        value = quoted.value;
        pos = quoted.nextPos;
        line += quoted.newlines;
      }
    }
    if (value === null) {
      const start = pos;
      while (pos < n && !"\t\r\n".includes(text[pos])) pos++;
      value = text.slice(start, pos);
    }
    fields.push(value);

    if (pos >= n) {
      rows.push({ line: rowLine, fields });
      break;
    }
    if (text[pos] === "\t") {
      pos++;
      continue;
    }

    // quebra de linha
    if (text[pos] === "\r" && pos + 1 < n && text[pos + 1] === "\n") pos += 2;
    else pos++;
    rows.push({ line: rowLine, fields });
    fields = [];
    line++;
    rowLine = line;
    if (pos >= n) break;
  }
  return rows;
}

function isBlank(row: RawRow): boolean {
  return row.fields.every((f) => !f.trim());
}

function trimBlankEdges(rows: RawRow[]): RawRow[] {
  let start = 0;
  let end = rows.length;
  while (start < end && isBlank(rows[start])) start++;
  while (end > start && isBlank(rows[end - 1])) end--;
  return rows.slice(start, end);
}

/** Lê e valida o texto colado. Nunca descarta linhas em silêncio. */
export function parsePromotions(text: string): ParseResult {
  const result: ParseResult = { records: [], errors: [], warnings: [], rowsRead: 0 };
  const rows = trimBlankEdges(splitTsv(text));
  result.rowsRead = rows.length;

  if (rows.length === 0) {
    result.errors.push({
      line: 1,
      message: "nenhum dado encontrado. Cole o resultado da consulta SQL Server na área de texto.",
      severity: "erro",
    });
    return result;
  }

  const validWithLines: Array<readonly [number, PromoRecord]> = [];
  for (const row of rows) {
    if (isBlank(row)) {
      result.errors.push({
        line: row.line,
        message: "linha em branco no meio dos dados. Remova-a para evitar perda de registros.",
        severity: "erro",
      });
      continue;
    }

    if (looksLikeHeader(row.fields)) {
      const nomes = EXPECTED_HEADERS.join(", ");
      result.errors.push({
        line: row.line,
        message:
          `parece a linha de cabeçalhos (${nomes}) colada junto com os dados. ` +
          "Apague essa linha, ou copie o resultado sem cabeçalhos " +
          '(no SSMS use Ctrl+C na grade, e não "Copiar com cabeçalhos"). ' +
          "Os cabeçalhos já existem na planilha modelo.",
        severity: "erro",
      });
      continue;
    }

    const { record, errors, warnings } = validateRow(row);
    result.errors.push(...errors);
    result.warnings.push(...warnings);
    if (record) {
      result.records.push(record);
      validWithLines.push([row.line, record]);
    }
  }

  result.warnings.push(...crossCheck(validWithLines));
  result.warnings.sort((a, b) => a.line - b.line);
  result.errors.sort((a, b) => a.line - b.line);
  return result;
}

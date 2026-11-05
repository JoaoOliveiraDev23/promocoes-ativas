/**
 * Validação e conversão dos campos de cada registro.
 *
 * Regras (decididas a partir da planilha real):
 *  - COD e PRIORIDADE são números inteiros; valores com zeros à esquerda são REJEITADOS
 *    (em vez de convertidos em silêncio), porque o zero se perderia.
 *  - DATA INICIAL e DATA FINAL são datas reais; a entrada deve estar estritamente em AAAA-MM-DD.
 *  - NOME e DETALHES são texto e são preservados exatamente como colados.
 *  - Apenas DETALHES pode ficar vazio.
 */

import {
  COL,
  COLUMN_COUNT,
  DATE_FORMAT_HUMAN,
  EXCEL_MAX_CELL_CHARS,
  EXPECTED_HEADERS,
  MAX_INTEGER_DIGITS,
  MIN_YEAR,
  NULL_TOKEN,
  OPTIONAL_COLUMNS,
} from "./constants";
import type { CalendarDate, Issue, PromoRecord, RawRow } from "./models";

/** Erro de validação de um campo; a mensagem já é amigável (pt-BR). */
export class FieldError extends Error {}

/**
 * Caracteres que o formato XLSX/XML não aceita: controles (exceto tab, LF, CR),
 * U+FFFE/U+FFFF e substitutos (surrogates) soltos.
 */
// eslint-disable-next-line no-control-regex
const ILLEGAL_XML_CHARS = /[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/;
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/;
const ISO_DATE_WITH_TIME = /^\d{4}-\d{2}-\d{2}[ T]\d/;
const BR_DATE = /^\d{1,2}\/\d{1,2}\/\d{2,4}$/;
const DIGITS_ONLY = /^[0-9]+$/;

/** Trecho do valor para mensagens de erro (curto, entre aspas). */
function shown(value: string, limit = 40): string {
  let v = value.replaceAll("\n", "↵");
  if (v.length > limit) v = v.slice(0, limit - 1) + "…";
  return `'${v}'`;
}

// --------------------------------------------------------------------------- //
// Cabeçalhos colados por engano
// --------------------------------------------------------------------------- //
export function normalizeHeaderName(value: string): string {
  return value.trim().toUpperCase().split(/\s+/).filter(Boolean).join(" ");
}

/**
 * True se a linha parece ser a linha de cabeçalhos (COD, NOME, ...).
 * Basta que pelo menos três nomes de cabeçalho esperados apareçam na linha.
 */
export function looksLikeHeader(fields: readonly string[]): boolean {
  const found = new Set(fields.map(normalizeHeaderName));
  let hits = 0;
  for (const header of EXPECTED_HEADERS) if (found.has(header)) hits++;
  return hits >= 3;
}

// --------------------------------------------------------------------------- //
// Conversores de campo
// --------------------------------------------------------------------------- //
/** Converte inteiro não negativo, sem zeros à esquerda. Lança FieldError. */
export function parseInteger(value: string): number {
  const text = value.trim();
  if (!text) throw new FieldError("campo obrigatório vazio.");
  if (!DIGITS_ONLY.test(text)) {
    throw new FieldError(`esperado número inteiro (somente dígitos), encontrado ${shown(value)}.`);
  }
  if (text.length > 1 && text.startsWith("0")) {
    throw new FieldError(
      `valor ${shown(value)} tem zeros à esquerda; na planilha esta coluna é numérica ` +
        "e o zero seria perdido. Corrija a consulta/dado de origem.",
    );
  }
  if (text.length > MAX_INTEGER_DIGITS) {
    throw new FieldError(`número grande demais (${text.length} dígitos; máximo ${MAX_INTEGER_DIGITS}).`);
  }
  return Number(text);
}

function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

/** Converte AAAA-MM-DD em data de calendário. Lança FieldError. */
export function parseIsoDate(value: string): CalendarDate {
  const text = value.trim();
  if (!text) throw new FieldError("campo obrigatório vazio.");
  if (ISO_DATE_WITH_TIME.test(text)) {
    throw new FieldError(
      `data com horário (${shown(value)}); esperado somente ${DATE_FORMAT_HUMAN}. ` +
        "Na consulta, converta a data (ex.: CONVERT(varchar(10), campo, 23)).",
    );
  }
  if (BR_DATE.test(text)) {
    throw new FieldError(`formato dia/mês/ano não aceito (${shown(value)}); esperado ${DATE_FORMAT_HUMAN}.`);
  }
  const match = ISO_DATE.exec(text);
  if (!match) throw new FieldError(`data inválida ${shown(value)}; esperado ${DATE_FORMAT_HUMAN}.`);
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (year < MIN_YEAR) {
    throw new FieldError(`ano ${year} anterior a ${MIN_YEAR}, que o Excel não representa.`);
  }
  if (month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month)) {
    throw new FieldError(`data inexistente ${shown(value)}.`);
  }
  return { year, month, day };
}

function checkText(value: string, required: boolean): void {
  if (required && !value.trim()) throw new FieldError("campo obrigatório vazio.");
  if (ILLEGAL_XML_CHARS.test(value)) throw new FieldError("contém caractere de controle inválido para o Excel.");
  if (value.length > EXCEL_MAX_CELL_CHARS) {
    throw new FieldError(
      `texto com ${value.length} caracteres; o Excel aceita no máximo ${EXCEL_MAX_CELL_CHARS}.`,
    );
  }
}

// --------------------------------------------------------------------------- //
// Validação de uma linha
// --------------------------------------------------------------------------- //
export interface RowValidation {
  record: PromoRecord | null;
  errors: Issue[];
  warnings: Issue[];
}

/**
 * Valida uma linha. Se houver qualquer erro, o registro é null e a linha nunca é
 * descartada em silêncio: todos os problemas dela são relatados.
 */
export function validateRow(row: RawRow): RowValidation {
  const errors: Issue[] = [];
  const warnings: Issue[] = [];

  if (row.fields.length !== COLUMN_COUNT) {
    const found = row.fields.length;
    let hint: string;
    if (found === 1) {
      hint = " Nenhuma tabulação encontrada: confira se o texto foi copiado direto da grade de resultados do SSMS.";
    } else if (found > COLUMN_COUNT) {
      hint = " Algum texto pode conter tabulação ou há colunas a mais na consulta.";
    } else {
      hint = " Pode haver coluna faltando ou uma quebra de linha dentro de um campo de texto.";
    }
    errors.push({
      line: row.line,
      message: `esperadas ${COLUMN_COUNT} colunas, encontradas ${found}.${hint}`,
      severity: "erro",
    });
    return { record: null, errors, warnings };
  }

  const fields = [...row.fields];

  // Tratamento do texto "NULL" (como o SSMS copia valores nulos).
  fields.forEach((value, idx) => {
    if (value.trim().toUpperCase() !== NULL_TOKEN) return;
    const name = EXPECTED_HEADERS[idx];
    if (OPTIONAL_COLUMNS.has(idx)) {
      fields[idx] = "";
      warnings.push({ line: row.line, message: "valor NULL tratado como vazio.", column: name, severity: "aviso" });
    } else {
      errors.push({ line: row.line, message: "valor NULL em campo obrigatório.", column: name, severity: "erro" });
    }
  });

  /** Aplica o conversor; registra erro (com o nome da coluna) em caso de falha. */
  function run<T>(idx: number, converter: (v: string) => T): T | undefined {
    if (errors.some((e) => e.column === EXPECTED_HEADERS[idx])) return undefined; // já reportado como NULL
    try {
      return converter(fields[idx]);
    } catch (exc) {
      if (exc instanceof FieldError) {
        errors.push({ line: row.line, message: exc.message, column: EXPECTED_HEADERS[idx], severity: "erro" });
        return undefined;
      }
      throw exc;
    }
  }

  const cod = run(COL.COD, parseInteger);
  const nome = run(COL.NOME, (v) => (checkText(v, true), v));
  const detalhes = run(COL.DETALHES, (v) => (checkText(v, !OPTIONAL_COLUMNS.has(COL.DETALHES)), v));
  const dataInicial = run(COL.DATA_INICIAL, parseIsoDate);
  const dataFinal = run(COL.DATA_FINAL, parseIsoDate);
  const prioridade = run(COL.PRIORIDADE, parseInteger);

  if (errors.length > 0) return { record: null, errors, warnings };

  return {
    record: {
      cod: cod!,
      nome: nome!,
      detalhes: detalhes!,
      dataInicial: dataInicial!,
      dataFinal: dataFinal!,
      prioridade: prioridade!,
    },
    errors,
    warnings,
  };
}

// --------------------------------------------------------------------------- //
// Verificações entre registros (somente avisos; não bloqueiam a geração)
// --------------------------------------------------------------------------- //
/**
 * Avisos sobre COD repetido e DATA FINAL anterior à DATA INICIAL.
 * São apenas avisos: a fonte oficial é a consulta e nenhum registro é filtrado.
 */
export function crossCheck(recordsWithLines: ReadonlyArray<readonly [number, PromoRecord]>): Issue[] {
  const warnings: Issue[] = [];
  const seen = new Map<number, number[]>();
  const key = (d: CalendarDate) => d.year * 10000 + d.month * 100 + d.day;

  for (const [line, rec] of recordsWithLines) {
    const lines = seen.get(rec.cod) ?? [];
    lines.push(line);
    seen.set(rec.cod, lines);
    if (key(rec.dataFinal) < key(rec.dataInicial)) {
      warnings.push({ line, message: "DATA FINAL é anterior à DATA INICIAL.", column: "DATA FINAL", severity: "aviso" });
    }
  }

  for (const [cod, lines] of seen) {
    if (lines.length > 1) {
      warnings.push({
        line: lines[1],
        message: `COD ${cod} aparece mais de uma vez (linhas ${lines.join(", ")}).`,
        column: "COD",
        severity: "aviso",
      });
    }
  }
  warnings.sort((a, b) => a.line - b.line);
  return warnings;
}

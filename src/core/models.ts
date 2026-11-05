/** Estruturas de dados compartilhadas entre parser, validadores e Excel. */

/** Linha bruta extraída do texto colado, já separada em campos. */
export interface RawRow {
  /** Número da (primeira) linha física no texto colado, base 1. */
  line: number;
  fields: string[];
}

/** Data de calendário sem fuso horário (evita surpresas do objeto Date). */
export interface CalendarDate {
  year: number;
  month: number; // 1..12
  day: number;
}

/** Um registro de promoção já validado e convertido. */
export interface PromoRecord {
  cod: number;
  nome: string;
  detalhes: string; // "" quando vazio
  dataInicial: CalendarDate;
  dataFinal: CalendarDate;
  prioridade: number;
}

/** Problema encontrado na validação (erro bloqueante ou aviso). */
export interface Issue {
  line: number;
  message: string;
  column?: string;
  severity: "erro" | "aviso";
}

export function describeIssue(issue: Issue): string {
  const coluna = issue.column ? ` [${issue.column}]` : "";
  return `Linha ${issue.line}${coluna}: ${issue.message}`;
}

/** Resultado completo da leitura + validação do texto colado. */
export interface ParseResult {
  records: PromoRecord[];
  errors: Issue[];
  warnings: Issue[];
  /** Linhas de dados lidas (após remover brancas nas pontas). */
  rowsRead: number;
}

/** Só é válido se não há erros e existe ao menos um registro. */
export function isValid(result: ParseResult): boolean {
  return result.errors.length === 0 && result.records.length > 0;
}

export function dateKey(d: CalendarDate): number {
  return d.year * 10000 + d.month * 100 + d.day;
}

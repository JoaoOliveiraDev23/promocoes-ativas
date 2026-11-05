/**
 * Geração em uma única chamada: texto colado -> planilha atualizada.
 * Toda a validação ocorre aqui, de forma automática, antes de gerar.
 * O modelo (bytes) é apenas LIDO; o resultado é devolvido em memória.
 */

import { MAX_ERRORS_SHOWN, TEMPLATE_DISPLAY_PATH } from "./constants";
import { parsePromotions } from "./dataParser";
import { buildOutputFilename } from "./filename";
import { describeIssue, isValid } from "./models";
import { ExcelServiceError, TemplateError, updateWorkbook } from "./xlsx";

/** Os dados colados têm problemas que impedem a geração. */
export class InvalidInput extends ExcelServiceError {
  constructor(
    public readonly lines: string[],
    public readonly total: number = lines.length,
  ) {
    super(lines.join("\n"));
  }
}

export interface GeneratedFile {
  filename: string;
  data: Uint8Array;
  recordCount: number;
  warnings: string[];
}

/** Valida o texto e gera a planilha. Lança ExcelServiceError (mensagem amigável). */
export function generateWorkbook(text: string, template: Uint8Array, now: Date = new Date()): GeneratedFile {
  if (!text.trim()) throw new InvalidInput(["Cole o resultado da consulta SQL Server no campo de texto."]);

  const result = parsePromotions(text);
  if (!isValid(result)) {
    throw new InvalidInput(result.errors.slice(0, MAX_ERRORS_SHOWN).map(describeIssue), result.errors.length);
  }

  let data: Uint8Array;
  try {
    data = updateWorkbook(template, result.records);
  } catch (exc) {
    if (exc instanceof TemplateError) {
      throw new TemplateError(`O modelo interno (${TEMPLATE_DISPLAY_PATH}) não é válido: ${exc.message}`);
    }
    throw exc;
  }

  return {
    filename: buildOutputFilename(now),
    data,
    recordCount: result.records.length,
    warnings: result.warnings.slice(0, MAX_ERRORS_SHOWN).map(describeIssue),
  };
}

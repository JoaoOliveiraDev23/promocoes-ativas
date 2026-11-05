/** Erro base; a mensagem é amigável e pode ser mostrada ao usuário. */
export class ExcelServiceError extends Error {}

/** O arquivo de modelo não é compatível (ou não foi encontrado). */
export class TemplateError extends ExcelServiceError {}

/** Falha ao gerar ou conferir o arquivo de saída. */
export class GenerationError extends ExcelServiceError {}

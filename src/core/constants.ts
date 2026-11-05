/**
 * Constantes do modelo oficial da planilha de promoções.
 * (Decididas a partir da inspeção da planilha real: aba "Planilha1", tabela "Tabela1".)
 */

/** Cabeçalhos da linha 1, na ordem exata das colunas A..F. */
export const EXPECTED_HEADERS = [
  "COD",
  "NOME",
  "DETALHES",
  "DATA INICIAL",
  "DATA FINAL",
  "PRIORIDADE",
] as const;

export const COLUMN_COUNT = EXPECTED_HEADERS.length;

/** Índices (base 0) das colunas dentro de cada registro. */
export const COL = {
  COD: 0,
  NOME: 1,
  DETALHES: 2,
  DATA_INICIAL: 3,
  DATA_FINAL: 4,
  PRIORIDADE: 5,
} as const;

/** Primeira linha de dados na planilha (a linha 1 guarda os cabeçalhos). */
export const FIRST_DATA_ROW = 2;

/** Prefixo do nome do arquivo de saída. */
export const OUTPUT_PREFIX = "PROMOCOES-ATIVAS";

/** Formato esperado para as datas coladas (resultado do SQL Server). */
export const DATE_FORMAT_HUMAN = "AAAA-MM-DD";

export const EXCEL_MAX_CELL_CHARS = 32767; // limite do Excel por célula
export const MAX_INTEGER_DIGITS = 15; // acima disso o Excel perde precisão
export const MIN_YEAR = 1900; // o Excel não representa datas anteriores a 1900

/**
 * Colunas que podem ficar vazias. Na planilha real, COD/NOME/DATAS/PRIORIDADE estão
 * preenchidos em todos os registros; DETALHES é texto livre e é a única opcional.
 */
export const OPTIONAL_COLUMNS: ReadonlySet<number> = new Set([COL.DETALHES]);

/** O SSMS copia valores NULL como o texto "NULL". */
export const NULL_TOKEN = "NULL";

/** Quantidade máxima de erros listados na tela. */
export const MAX_ERRORS_SHOWN = 5;

/** Caminho (relativo à raiz do site) do modelo interno. */
export const TEMPLATE_FILE_NAME = "PROMOCOES-ATIVAS-modelo.xlsx";
export const TEMPLATE_DISPLAY_PATH = `public/${TEMPLATE_FILE_NAME}`;

"""Constantes do modelo oficial da planilha de promoções.

Todas as decisões abaixo foram tiradas da inspeção do arquivo real
``PROMOCOES-ATIVAS DD-MM-AAAA.xlsx`` (aba ``Planilha1``, tabela ``Tabela1``).
"""

from __future__ import annotations

# Cabeçalhos da linha 1, na ordem exata das colunas A..F.
EXPECTED_HEADERS: tuple[str, ...] = (
    "COD",
    "NOME",
    "DETALHES",
    "DATA INICIAL",
    "DATA FINAL",
    "PRIORIDADE",
)

COLUMN_COUNT: int = len(EXPECTED_HEADERS)

# Índices (0-based) das colunas dentro de cada registro.
COL_COD, COL_NOME, COL_DETALHES, COL_DATA_INICIAL, COL_DATA_FINAL, COL_PRIORIDADE = range(6)

# Primeira linha de dados na planilha (a linha 1 guarda os cabeçalhos).
FIRST_DATA_ROW: int = 2

# Prefixo do nome do arquivo de saída.
OUTPUT_PREFIX: str = "PROMOCOES-ATIVAS"

# Formato esperado para as datas coladas (resultado do SQL Server).
DATE_FORMAT_HUMAN: str = "AAAA-MM-DD"

# Limites técnicos.
EXCEL_MAX_CELL_CHARS: int = 32767  # limite do Excel por célula
MAX_INTEGER_DIGITS: int = 15  # acima disso o Excel perde precisão
MIN_YEAR: int = 1900  # o Excel não representa datas anteriores a 1900

# Colunas que podem ficar vazias. Na planilha real, COD/NOME/DATAS/PRIORIDADE
# estão preenchidos em todos os 98 registros; DETALHES é texto livre e é a
# única coluna tratada como opcional.
OPTIONAL_COLUMNS: frozenset[int] = frozenset({COL_DETALHES})

# SSMS copia valores NULL como o texto "NULL".
NULL_TOKEN: str = "NULL"

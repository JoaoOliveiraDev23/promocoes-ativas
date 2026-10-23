"""Validação e conversão dos campos de cada registro.

Regras (decididas a partir da planilha real):

* ``COD`` e ``PRIORIDADE`` são números inteiros na planilha (98 de 98 registros),
  portanto são gravados como número. Valores com zeros à esquerda são
  *rejeitados* (em vez de convertidos em silêncio), porque o zero se perderia.
* ``DATA INICIAL`` e ``DATA FINAL`` são datas reais (``datetime`` à meia-noite);
  a entrada deve estar estritamente em ``AAAA-MM-DD``.
* ``NOME`` e ``DETALHES`` são texto e são preservados exatamente como colados
  (inclusive espaços nas pontas e espaços duplos, que existem nos dados reais).
* Apenas ``DETALHES`` pode ficar vazio.
"""

from __future__ import annotations

import re
from collections import defaultdict
from datetime import date

from .constants import (
    COL_COD,
    COL_DATA_FINAL,
    COL_DATA_INICIAL,
    COL_DETALHES,
    COL_NOME,
    COL_PRIORIDADE,
    COLUMN_COUNT,
    DATE_FORMAT_HUMAN,
    EXCEL_MAX_CELL_CHARS,
    EXPECTED_HEADERS,
    MAX_INTEGER_DIGITS,
    MIN_YEAR,
    NULL_TOKEN,
    OPTIONAL_COLUMNS,
)
from .models import Issue, PromoRecord, RawRow

# Caracteres que o formato XLSX não aceita (mesmo conjunto rejeitado pelo openpyxl).
_ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_ISO_DATE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
_ISO_DATE_WITH_TIME = re.compile(r"\d{4}-\d{2}-\d{2}[ T]\d")
_BR_DATE = re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}")
_DIGITS_ONLY = re.compile(r"[0-9]+")


class FieldError(ValueError):
    """Erro de validação de um campo; a mensagem já é amigável (pt-BR)."""


def _shown(value: str, limit: int = 40) -> str:
    """Trecho do valor para mensagens de erro (curto, entre aspas)."""
    value = value.replace("\n", "↵")
    if len(value) > limit:
        value = value[: limit - 1] + "…"
    return f"'{value}'"


# --------------------------------------------------------------------------- #
# Cabeçalhos colados por engano
# --------------------------------------------------------------------------- #
def normalize_header_name(value: str) -> str:
    return " ".join(value.strip().upper().split())


def looks_like_header(fields: tuple[str, ...] | list[str]) -> bool:
    """True se a linha parece ser a linha de cabeçalhos (COD, NOME, ...).

    Basta que pelo menos três nomes de cabeçalho esperados apareçam na linha;
    assim também pegamos cabeçalhos colados fora de ordem ou com uma coluna a mais.
    """
    found = {normalize_header_name(f) for f in fields}
    return len(found.intersection(EXPECTED_HEADERS)) >= 3


# --------------------------------------------------------------------------- #
# Conversores de campo
# --------------------------------------------------------------------------- #
def parse_integer(value: str) -> int:
    """Converte inteiro não negativo, sem zeros à esquerda. Levanta FieldError."""
    text = value.strip()
    if not text:
        raise FieldError("campo obrigatório vazio.")
    if not _DIGITS_ONLY.fullmatch(text):
        raise FieldError(f"esperado número inteiro (somente dígitos), encontrado {_shown(value)}.")
    if len(text) > 1 and text.startswith("0"):
        raise FieldError(
            f"valor {_shown(value)} tem zeros à esquerda; na planilha esta coluna é numérica "
            "e o zero seria perdido. Corrija a consulta/dado de origem."
        )
    if len(text) > MAX_INTEGER_DIGITS:
        raise FieldError(f"número grande demais ({len(text)} dígitos; máximo {MAX_INTEGER_DIGITS}).")
    return int(text)


def parse_iso_date(value: str) -> date:
    """Converte ``AAAA-MM-DD`` em ``date``. Levanta FieldError."""
    text = value.strip()
    if not text:
        raise FieldError("campo obrigatório vazio.")
    if _ISO_DATE_WITH_TIME.match(text):
        raise FieldError(
            f"data com horário ({_shown(value)}); esperado somente {DATE_FORMAT_HUMAN}. "
            "Na consulta, converta a data (ex.: CONVERT(varchar(10), campo, 23))."
        )
    if _BR_DATE.fullmatch(text):
        raise FieldError(f"formato dia/mês/ano não aceito ({_shown(value)}); esperado {DATE_FORMAT_HUMAN}.")
    match = _ISO_DATE.fullmatch(text)
    if not match:
        raise FieldError(f"data inválida {_shown(value)}; esperado {DATE_FORMAT_HUMAN}.")
    year, month, day = (int(g) for g in match.groups())
    if year < MIN_YEAR:
        raise FieldError(f"ano {year} anterior a {MIN_YEAR}, que o Excel não representa.")
    try:
        return date(year, month, day)
    except ValueError:
        raise FieldError(f"data inexistente {_shown(value)}.") from None


def _check_text(value: str, required: bool) -> None:
    if required and not value.strip():
        raise FieldError("campo obrigatório vazio.")
    if _ILLEGAL_XML_CHARS.search(value):
        raise FieldError("contém caractere de controle inválido para o Excel.")
    if len(value) > EXCEL_MAX_CELL_CHARS:
        raise FieldError(f"texto com {len(value)} caracteres; o Excel aceita no máximo {EXCEL_MAX_CELL_CHARS}.")


# --------------------------------------------------------------------------- #
# Validação de uma linha
# --------------------------------------------------------------------------- #
def validate_row(row: RawRow) -> tuple[PromoRecord | None, list[Issue], list[Issue]]:
    """Valida uma linha. Retorna (registro | None, erros, avisos).

    Se houver qualquer erro, o registro é ``None`` e a linha nunca é descartada
    em silêncio: todos os problemas dela são relatados.
    """
    errors: list[Issue] = []
    warnings: list[Issue] = []

    if len(row.fields) != COLUMN_COUNT:
        found = len(row.fields)
        hint = ""
        if found == 1:
            hint = " Nenhuma tabulação encontrada: confira se o texto foi copiado direto da grade de resultados do SSMS."
        elif found > COLUMN_COUNT:
            hint = " Algum texto pode conter tabulação ou há colunas a mais na consulta."
        else:
            hint = " Pode haver coluna faltando ou uma quebra de linha dentro de um campo de texto."
        errors.append(Issue(row.line, f"esperadas {COLUMN_COUNT} colunas, encontradas {found}.{hint}"))
        return None, errors, warnings

    fields = list(row.fields)

    # Tratamento do texto "NULL" (como o SSMS copia valores nulos).
    for idx, value in enumerate(fields):
        if value.strip().upper() != NULL_TOKEN:
            continue
        name = EXPECTED_HEADERS[idx]
        if idx in OPTIONAL_COLUMNS:
            fields[idx] = ""
            warnings.append(Issue(row.line, "valor NULL tratado como vazio.", name, "aviso"))
        else:
            errors.append(Issue(row.line, "valor NULL em campo obrigatório.", name))

    def run(idx: int, converter):
        """Aplica o conversor; registra erro (com nome da coluna) em caso de falha."""
        if any(e.column == EXPECTED_HEADERS[idx] for e in errors):
            return None  # já reportado como NULL
        try:
            return converter(fields[idx])
        except FieldError as exc:
            errors.append(Issue(row.line, str(exc), EXPECTED_HEADERS[idx]))
            return None

    cod = run(COL_COD, parse_integer)
    nome = run(COL_NOME, lambda v: (_check_text(v, required=True), v)[1])
    detalhes = run(
        COL_DETALHES,
        lambda v: (_check_text(v, required=COL_DETALHES not in OPTIONAL_COLUMNS), v)[1],
    )
    data_inicial = run(COL_DATA_INICIAL, parse_iso_date)
    data_final = run(COL_DATA_FINAL, parse_iso_date)
    prioridade = run(COL_PRIORIDADE, parse_integer)

    if errors:
        return None, errors, warnings

    record = PromoRecord(cod, nome, detalhes, data_inicial, data_final, prioridade)
    return record, errors, warnings


# --------------------------------------------------------------------------- #
# Verificações entre registros (somente avisos; não bloqueiam a geração)
# --------------------------------------------------------------------------- #
def cross_check(records_with_lines: list[tuple[int, PromoRecord]]) -> list[Issue]:
    """Avisos sobre COD repetido e DATA FINAL anterior à DATA INICIAL.

    São apenas avisos: a fonte oficial é a consulta e nenhum registro é filtrado.
    """
    warnings: list[Issue] = []
    seen: dict[int, list[int]] = defaultdict(list)

    for line, rec in records_with_lines:
        seen[rec.cod].append(line)
        if rec.data_final < rec.data_inicial:
            warnings.append(
                Issue(line, "DATA FINAL é anterior à DATA INICIAL.", "DATA FINAL", "aviso")
            )

    for cod, lines in seen.items():
        if len(lines) > 1:
            where = ", ".join(str(n) for n in lines)
            warnings.append(
                Issue(lines[1], f"COD {cod} aparece mais de uma vez (linhas {where}).", "COD", "aviso")
            )
    warnings.sort(key=lambda w: w.line)
    return warnings

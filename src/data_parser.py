"""Leitura do texto colado do SQL Server (colunas separadas por tabulação)."""

from __future__ import annotations

from .constants import COLUMN_COUNT, EXPECTED_HEADERS
from .models import Issue, ParseResult, RawRow
from .validators import cross_check, looks_like_header, validate_row

_NEWLINES = "\r\n"


def _read_quoted(text: str, pos: int) -> tuple[str, int, int] | None:
    """Tenta ler um campo entre aspas (como o Excel faz ao colar texto).

    Só é considerado "entre aspas" se a aspa de fechamento for seguida de
    tabulação, quebra de linha ou fim do texto. Caso contrário o campo é lido
    literalmente (uma aspa solta no meio de um texto não altera nada).

    Retorna ``(valor, nova_posicao, quebras_de_linha_consumidas)`` ou ``None``.
    """
    n = len(text)
    i = pos + 1
    out: list[str] = []
    newlines = 0
    while i < n:
        ch = text[i]
        if ch == '"':
            if i + 1 < n and text[i + 1] == '"':  # aspa escapada ("")
                out.append('"')
                i += 2
                continue
            end = i + 1
            if end >= n or text[end] in "\t\r\n":
                value = "".join(out).replace("\r\n", "\n").replace("\r", "\n")
                return value, end, newlines
            return None
        if ch == "\n":
            newlines += 1
        elif ch == "\r" and not (i + 1 < n and text[i + 1] == "\n"):
            newlines += 1  # \r isolado; em \r\n a contagem é feita no \n
        out.append(ch)
        i += 1
    return None


def split_tsv(text: str) -> list[RawRow]:
    """Divide o texto em linhas e campos (separador: tabulação).

    * Aceita quebras ``\\n``, ``\\r\\n`` e ``\\r``.
    * Não remove espaços dos campos.
    * Não cria linha extra quando o texto termina com quebra de linha.
    * ``line`` é o número da primeira linha física do registro (base 1).
    """
    if text.startswith("\ufeff"):
        text = text[1:]

    rows: list[RawRow] = []
    n = len(text)
    pos = 0
    line = 1
    row_line = 1
    fields: list[str] = []

    while True:
        value: str | None = None
        if pos < n and text[pos] == '"':
            quoted = _read_quoted(text, pos)
            if quoted is not None:
                value, pos, consumed = quoted
                line += consumed
        if value is None:
            start = pos
            while pos < n and text[pos] not in "\t" + _NEWLINES:
                pos += 1
            value = text[start:pos]
        fields.append(value)

        if pos >= n:
            rows.append(RawRow(row_line, tuple(fields)))
            break
        if text[pos] == "\t":
            pos += 1
            continue

        # quebra de linha
        if text[pos] == "\r" and pos + 1 < n and text[pos + 1] == "\n":
            pos += 2
        else:
            pos += 1
        rows.append(RawRow(row_line, tuple(fields)))
        fields = []
        line += 1
        row_line = line
        if pos >= n:
            break

    return rows


def _is_blank(row: RawRow) -> bool:
    return all(not f.strip() for f in row.fields)


def _trim_blank_edges(rows: list[RawRow]) -> list[RawRow]:
    start, end = 0, len(rows)
    while start < end and _is_blank(rows[start]):
        start += 1
    while end > start and _is_blank(rows[end - 1]):
        end -= 1
    return rows[start:end]


def parse_promotions(text: str) -> ParseResult:
    """Lê e valida o texto colado. Nunca descarta linhas em silêncio."""
    result = ParseResult()
    rows = _trim_blank_edges(split_tsv(text))
    result.rows_read = len(rows)

    if not rows:
        result.errors.append(
            Issue(1, "nenhum dado encontrado. Cole o resultado da consulta SQL Server na área de texto.")
        )
        return result

    valid_with_lines = []
    for row in rows:
        if _is_blank(row):
            result.errors.append(
                Issue(row.line, "linha em branco no meio dos dados. Remova-a para evitar perda de registros.")
            )
            continue

        if looks_like_header(row.fields):
            nomes = ", ".join(EXPECTED_HEADERS)
            result.errors.append(
                Issue(
                    row.line,
                    f"parece a linha de cabeçalhos ({nomes}) colada junto com os dados. "
                    "Apague essa linha, ou copie o resultado sem cabeçalhos "
                    "(no SSMS use Ctrl+C na grade, e não \"Copiar com cabeçalhos\"). "
                    "Os cabeçalhos já existem na planilha modelo.",
                )
            )
            continue

        record, errors, warnings = validate_row(row)
        result.errors.extend(errors)
        result.warnings.extend(warnings)
        if record is not None:
            result.records.append(record)
            valid_with_lines.append((row.line, record))

    result.warnings.extend(cross_check(valid_with_lines))
    result.warnings.sort(key=lambda w: w.line)
    result.errors.sort(key=lambda e: e.line)
    return result


__all__ = ["split_tsv", "parse_promotions", "COLUMN_COUNT"]

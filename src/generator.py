"""Geração em uma única chamada: texto colado -> planilha atualizada.

Usa o modelo interno (``assets/PROMOCOES-ATIVAS-modelo.xlsx``), que só é LIDO:
nada é gravado no modelo nem em disco; o resultado é devolvido em memória.
Toda a validação ocorre aqui, de forma automática, antes de gerar.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from .data_parser import parse_promotions
from .excel_service import ExcelServiceError, TemplateError, build_output_filename, update_workbook
from .models import Issue

TEMPLATE_RELATIVE_PATH = "assets/PROMOCOES-ATIVAS-modelo.xlsx"
DEFAULT_TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "assets" / "PROMOCOES-ATIVAS-modelo.xlsx"
MAX_ERRORS_SHOWN = 5


class InvalidInput(ExcelServiceError):
    """Os dados colados têm problemas que impedem a geração."""

    def __init__(self, lines: list[str], total: int | None = None):
        self.lines = lines
        self.total = total if total is not None else len(lines)
        super().__init__("\n".join(lines))


@dataclass
class GeneratedFile:
    filename: str
    data: bytes
    record_count: int
    warnings: list[str] = field(default_factory=list)


def load_template(path: str | Path | None = None) -> bytes:
    """Lê o modelo interno (somente leitura). Mensagem clara se estiver ausente."""
    template_path = Path(path) if path is not None else DEFAULT_TEMPLATE_PATH
    if not template_path.is_file():
        raise TemplateError(
            f"O arquivo de modelo interno não foi encontrado: {TEMPLATE_RELATIVE_PATH}. "
            "Restaure-o na pasta do projeto e tente novamente."
        )
    try:
        return template_path.read_bytes()
    except OSError as exc:
        raise TemplateError(
            f"Não foi possível ler o modelo interno ({TEMPLATE_RELATIVE_PATH}): {exc.strerror or type(exc).__name__}."
        ) from exc


def generate_workbook(
    text: str,
    *,
    template_path: str | Path | None = None,
    now: date | datetime | None = None,
) -> GeneratedFile:
    """Valida o texto e gera a planilha. Levanta ``ExcelServiceError`` (mensagem amigável)."""
    template = load_template(template_path)

    if not text.strip():
        raise InvalidInput(["Cole o resultado da consulta SQL Server no campo de texto."])

    result = parse_promotions(text)
    if not result.is_valid:
        shown = [issue.describe() for issue in result.errors[:MAX_ERRORS_SHOWN]]
        raise InvalidInput(shown, total=len(result.errors))

    try:
        data = update_workbook(template, result.records)
    except TemplateError as exc:
        raise TemplateError(f"O modelo interno ({TEMPLATE_RELATIVE_PATH}) não é válido: {exc}") from exc

    return GeneratedFile(
        filename=build_output_filename(now),
        data=data,
        record_count=len(result.records),
        warnings=[w.describe() for w in result.warnings[:MAX_ERRORS_SHOWN]],
    )


__all__ = [
    "DEFAULT_TEMPLATE_PATH",
    "TEMPLATE_RELATIVE_PATH",
    "GeneratedFile",
    "InvalidInput",
    "Issue",
    "generate_workbook",
    "load_template",
]

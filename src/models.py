"""Estruturas de dados compartilhadas entre parser, validadores e Excel."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Literal


@dataclass(frozen=True)
class RawRow:
    """Linha bruta extraída do texto colado, já separada em campos."""

    line: int  # número da (primeira) linha física no texto colado, base 1
    fields: tuple[str, ...]


@dataclass(frozen=True)
class PromoRecord:
    """Um registro de promoção já validado e convertido."""

    cod: int
    nome: str
    detalhes: str  # "" quando vazio
    data_inicial: date
    data_final: date
    prioridade: int


@dataclass(frozen=True)
class Issue:
    """Problema encontrado na validação (erro bloqueante ou aviso)."""

    line: int
    message: str
    column: str | None = None
    severity: Literal["erro", "aviso"] = "erro"

    def describe(self) -> str:
        coluna = f" [{self.column}]" if self.column else ""
        return f"Linha {self.line}{coluna}: {self.message}"


@dataclass
class ParseResult:
    """Resultado completo da leitura + validação do texto colado."""

    records: list[PromoRecord] = field(default_factory=list)
    errors: list[Issue] = field(default_factory=list)
    warnings: list[Issue] = field(default_factory=list)
    rows_read: int = 0  # linhas de dados lidas (após remover brancas nas pontas)

    @property
    def is_valid(self) -> bool:
        """Só é válido se não há erros e existe ao menos um registro."""
        return not self.errors and bool(self.records)

from __future__ import annotations

from datetime import date

import pytest

from src.models import RawRow
from src.validators import (
    FieldError,
    looks_like_header,
    parse_integer,
    parse_iso_date,
    validate_row,
)


def row(*fields: str, line: int = 7) -> RawRow:
    return RawRow(line, tuple(fields))


GOOD = ("10", "Nome", "Detalhe", "2026-01-31", "2026-12-31", "2")


class TestDatas:
    def test_data_valida(self):
        assert parse_iso_date("2026-02-28") == date(2026, 2, 28)
        assert parse_iso_date(" 2024-02-29 ") == date(2024, 2, 29)  # ano bissexto

    @pytest.mark.parametrize(
        "valor, trecho",
        [
            ("", "obrigatório"),
            ("2026-02-30", "inexistente"),
            ("2025-02-29", "inexistente"),
            ("2026-13-01", "inexistente"),
            ("10/03/2025", "dia/mês/ano"),
            ("2026/03/10", "inválida"),
            ("26-03-10", "inválida"),
            ("2026-3-1", "inválida"),
            ("2026-03-10 00:00:00.000", "horário"),
            ("2026-03-10T10:00", "horário"),
            ("1899-12-31", "anterior"),
            ("abc", "inválida"),
        ],
    )
    def test_datas_invalidas(self, valor, trecho):
        with pytest.raises(FieldError, match=trecho):
            parse_iso_date(valor)


class TestInteiros:
    def test_inteiros_validos(self):
        assert parse_integer("187") == 187
        assert parse_integer(" 4 ") == 4
        assert parse_integer("0") == 0

    @pytest.mark.parametrize(
        "valor, trecho",
        [
            ("", "obrigatório"),
            ("1,5", "inteiro"),
            ("1.0", "inteiro"),
            ("-1", "inteiro"),
            ("abc", "inteiro"),
            ("1 2", "inteiro"),
            ("0123", "zeros à esquerda"),
            ("1" * 16, "grande demais"),
        ],
    )
    def test_inteiros_invalidos(self, valor, trecho):
        with pytest.raises(FieldError, match=trecho):
            parse_integer(valor)

    def test_digitos_nao_ascii_sao_rejeitados(self):
        with pytest.raises(FieldError):
            parse_integer("١٢٣")  # dígitos arábico-índicos


class TestCabecalho:
    def test_detecta_cabecalho_exato(self):
        assert looks_like_header(("COD", "NOME", "DETALHES", "DATA INICIAL", "DATA FINAL", "PRIORIDADE"))

    def test_ignora_caixa_e_espacos(self):
        assert looks_like_header((" cod ", "Nome", "detalhes", "data  inicial", "x", "y"))

    def test_registro_normal_nao_e_cabecalho(self):
        assert not looks_like_header(GOOD)

    def test_promocao_chamada_nome_nao_e_cabecalho(self):
        assert not looks_like_header(("5", "NOME", "COD", "2026-01-01", "2026-01-02", "1"))


class TestLinha:
    def test_linha_valida(self):
        rec, errors, warnings = validate_row(row(*GOOD))
        assert errors == [] and warnings == []
        assert (rec.cod, rec.nome, rec.prioridade) == (10, "Nome", 2)

    def test_erros_informam_linha_e_coluna(self):
        rec, errors, _ = validate_row(row("x", "Nome", "d", "31/01/2026", "2026-12-31", "alta", line=42))
        assert rec is None
        assert {(e.line, e.column) for e in errors} == {(42, "COD"), (42, "DATA INICIAL"), (42, "PRIORIDADE")}
        assert "Linha 42 [COD]" in errors[0].describe()

    def test_caractere_de_controle_e_rejeitado(self):
        _, errors, _ = validate_row(row("1", "No\x00me", "d", "2026-01-01", "2026-01-02", "1"))
        assert errors[0].column == "NOME" and "controle" in errors[0].message

    def test_texto_longo_demais(self):
        _, errors, _ = validate_row(row("1", "N", "x" * 32768, "2026-01-01", "2026-01-02", "1"))
        assert errors[0].column == "DETALHES" and "32767" in errors[0].message

    def test_nome_so_com_espacos_e_vazio(self):
        _, errors, _ = validate_row(row("1", "   ", "d", "2026-01-01", "2026-01-02", "1"))
        assert errors[0].column == "NOME"

    def test_formula_digitada_como_texto_e_aceita_como_texto(self):
        rec, errors, _ = validate_row(row("1", "=1+1", "=HYPERLINK(\"x\")", "2026-01-01", "2026-01-02", "1"))
        assert errors == [] and rec.nome == "=1+1"

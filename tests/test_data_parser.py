from __future__ import annotations

from datetime import date

from src.data_parser import parse_promotions, split_tsv

HEADER = "COD\tNOME\tDETALHES\tDATA INICIAL\tDATA FINAL\tPRIORIDADE"
ROW1 = "187\tPetit Deli Ciclo 3 - 2025\tGranéis até 12 Unidades - R$ 2,87\t2025-03-10\t2026-12-31\t4"
ROW2 = "233\tKIT VITALÍCIO VD 1ª COMPRA\tPrimeira compra\t2025-05-05\t2026-12-31\t1"


def test_importa_registros_validos_separados_por_tabulacao():
    result = parse_promotions(f"{ROW1}\n{ROW2}")
    assert result.is_valid
    assert result.errors == [] and result.warnings == []
    assert len(result.records) == 2
    first = result.records[0]
    assert (first.cod, first.prioridade) == (187, 4)
    assert first.data_inicial == date(2025, 3, 10)
    assert first.data_final == date(2026, 12, 31)


def test_acentos_e_caracteres_especiais_sao_preservados():
    texto = "Ç ª á ã é í ú – “aspas” R$ 1.249,74 (50%) & <b> ;,"
    result = parse_promotions(f"1\tNome {texto}\t{texto}\t2026-01-01\t2026-01-02\t1")
    assert result.is_valid
    assert result.records[0].detalhes == texto
    assert result.records[0].nome == f"Nome {texto}"


def test_espacos_internos_e_nas_pontas_do_texto_sao_preservados():
    result = parse_promotions("1\t  Nome  com  espaços \tDetalhe \t2026-01-01\t2026-01-02\t1")
    assert result.records[0].nome == "  Nome  com  espaços "
    assert result.records[0].detalhes == "Detalhe "


def test_campo_vazio_permitido_em_detalhes():
    result = parse_promotions("1\tNome\t\t2026-01-01\t2026-01-02\t1")
    assert result.is_valid
    assert result.records[0].detalhes == ""


def test_campos_obrigatorios_vazios_geram_erro_com_linha_e_coluna():
    result = parse_promotions(f"{ROW1}\n\t\tx\t\t2026-01-02\t")
    assert not result.is_valid
    colunas = {e.column for e in result.errors if e.line == 2}
    assert colunas == {"COD", "NOME", "DATA INICIAL", "PRIORIDADE"}


def test_quantidade_incorreta_de_colunas():
    result = parse_promotions(f"{ROW1}\n1\tNome\t2026-01-01\t2026-01-02\t1\n{ROW2}\n1\t2\t3\t4\t5\t6\t7")
    assert not result.is_valid
    assert [e.line for e in result.errors] == [2, 4]
    assert "esperadas 6 colunas, encontradas 5" in result.errors[0].message
    assert "encontradas 7" in result.errors[1].message
    assert len(result.records) == 2  # as linhas boas continuam identificadas


def test_texto_sem_tabulacao_explica_o_problema():
    result = parse_promotions("187 Petit Deli 2025-03-10 2026-12-31 4")
    assert not result.is_valid
    assert "Nenhuma tabulação" in result.errors[0].message


def test_cabecalho_colado_e_identificado_sem_corromper():
    result = parse_promotions(f"{HEADER}\n{ROW1}")
    assert not result.is_valid
    assert result.errors[0].line == 1
    assert "cabeçalhos" in result.errors[0].message
    assert "sem cabeçalhos" in result.errors[0].message
    assert len(result.records) == 1  # a linha de dados válida não é descartada


def test_cabecalho_fora_de_ordem_ou_em_minusculas_tambem_e_detectado():
    result = parse_promotions("nome\tcod\tdetalhes\tprioridade\tdata final\tdata inicial\n" + ROW1)
    assert not result.is_valid
    assert "cabeçalhos" in result.errors[0].message


def test_linhas_em_branco_no_inicio_e_no_fim_sao_ignoradas():
    result = parse_promotions(f"\n\n  \n{ROW1}\n{ROW2}\n\n\n   \n")
    assert result.is_valid and len(result.records) == 2
    assert result.rows_read == 2


def test_numero_de_linha_considera_as_linhas_em_branco_iniciais():
    result = parse_promotions(f"\n\n{ROW1}\n1\tx\t2026-01-01\t1")
    assert [e.line for e in result.errors] == [4]


def test_linha_em_branco_no_meio_e_erro():
    result = parse_promotions(f"{ROW1}\n\n{ROW2}")
    assert not result.is_valid
    assert result.errors[0].line == 2
    assert "em branco" in result.errors[0].message


def test_entrada_vazia_e_erro():
    for texto in ("", "\n\n", "   \t \n"):
        result = parse_promotions(texto)
        assert not result.is_valid
        assert "nenhum dado" in result.errors[0].message.lower()


def test_quebras_de_linha_windows_e_mac():
    for sep in ("\r\n", "\r"):
        result = parse_promotions(f"{ROW1}{sep}{ROW2}{sep}")
        assert result.is_valid and len(result.records) == 2


def test_bom_no_inicio_e_ignorado():
    assert parse_promotions("\ufeff" + ROW1).is_valid


def test_campo_entre_aspas_com_quebra_de_linha_como_no_excel():
    texto = f'1\tNome\t"linha 1\nlinha 2 com ""aspas"""\t2026-01-01\t2026-01-02\t1\n{ROW2}'
    result = parse_promotions(texto)
    assert result.is_valid
    assert result.records[0].detalhes == 'linha 1\nlinha 2 com "aspas"'
    assert len(result.records) == 2


def test_numeracao_de_linhas_apos_campo_multilinha():
    texto = '1\tNome\t"a\nb\nc"\t2026-01-01\t2026-01-02\t1\n1\tx\t2026-01-01\t1'
    result = parse_promotions(texto)
    assert [e.line for e in result.errors] == [4]


def test_aspa_solta_no_meio_do_texto_nao_altera_o_valor():
    result = parse_promotions('1\t"Natal" especial\tTem 5" de altura\t2026-01-01\t2026-01-02\t1')
    assert result.is_valid
    assert result.records[0].nome == '"Natal" especial'
    assert result.records[0].detalhes == 'Tem 5" de altura'


def test_null_em_detalhes_vira_vazio_com_aviso_e_em_obrigatorio_e_erro():
    ok = parse_promotions("1\tNome\tNULL\t2026-01-01\t2026-01-02\t1")
    assert ok.is_valid and ok.records[0].detalhes == ""
    assert ok.warnings and ok.warnings[0].severity == "aviso"
    ruim = parse_promotions("NULL\tNome\tx\t2026-01-01\t2026-01-02\t1")
    assert not ruim.is_valid
    assert ruim.errors[0].column == "COD" and "NULL" in ruim.errors[0].message


def test_avisos_cod_repetido_e_datas_invertidas_nao_bloqueiam():
    texto = "1\tA\tx\t2026-05-01\t2026-04-01\t1\n1\tB\tx\t2026-01-01\t2026-02-01\t1"
    result = parse_promotions(texto)
    assert result.is_valid
    mensagens = " | ".join(w.message for w in result.warnings)
    assert "anterior" in mensagens and "mais de uma vez" in mensagens


def test_split_tsv_nao_cria_linha_extra_no_fim():
    assert len(split_tsv("a\tb\nc\td\n")) == 2
    assert [r.fields for r in split_tsv("a\tb\n\nc")] == [("a", "b"), ("",), ("c",)]

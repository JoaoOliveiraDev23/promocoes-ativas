"""Interface mínima: um campo, um botão (roda o app sem navegador, via AppTest)."""

from __future__ import annotations

import base64
import re
from pathlib import Path

import pytest

AppTest = pytest.importorskip("streamlit.testing.v1").AppTest
APP = str(Path(__file__).resolve().parent.parent / "app.py")

TEXTO_VALIDO = "\n".join(
    f"{100 + i}\tPromoção {i}\tDetalhes {i}\t2026-01-{1 + i:02d}\t2026-12-31\t{1 + i % 3}" for i in range(7)
)


def novo_app() -> "AppTest":
    return AppTest.from_file(APP, default_timeout=60).run()


def test_pagina_tem_somente_um_campo_e_um_botao():
    at = novo_app()
    assert not at.exception
    assert len(at.text_area) == 1
    assert [b.label for b in at.button] == ["Gerar e baixar planilha"]
    for ausente in (at.file_uploader, at.selectbox, at.checkbox, at.radio, at.get("download_button")):
        assert len(ausente) == 0
    assert [m.value for m in at.markdown] == ["#### Atualizador de Promoções"]


def test_um_clique_valida_gera_e_inicia_o_download():
    at = novo_app()
    at.text_area[0].set_value(TEXTO_VALIDO)
    at.button[0].click().run()
    assert not at.exception and not at.error
    assert "7 registro(s)" in at.success[0].value
    srcdoc = at.get("iframe")[0].proto.srcdoc
    assert 'download="PROMOCOES-ATIVAS ' in srcdoc and ".xlsx" in srcdoc
    # o conteúdo embutido para download é um .xlsx (zip) válido
    payload = re.search(r'atob\("([^"]+)"\)', srcdoc).group(1)
    assert base64.b64decode(payload)[:2] == b"PK"


def test_erro_mostra_linha_e_motivo_e_nao_inicia_download():
    at = novo_app()
    at.text_area[0].set_value(TEXTO_VALIDO + "\nabc\tNome\t\t10/03/2025\t2026-12-31\t1")
    at.button[0].click().run()
    assert not at.exception and not at.success
    mensagem = at.error[0].value
    assert "Linha 8" in mensagem and "COD" in mensagem
    assert len(at.get("iframe")) == 0


def test_campo_vazio_pede_para_colar_os_dados():
    at = novo_app()
    at.button[0].click().run()
    assert "Cole o resultado" in at.error[0].value and len(at.get("iframe")) == 0


def test_cifrao_e_underscore_dos_dados_nao_quebram_a_mensagem_de_erro():
    at = novo_app()
    at.text_area[0].set_value("1\tNome R$ 2,87 _x_\tD\t2026-01-01\n")
    at.button[0].click().run()
    assert not at.exception and "R$ 2,87" not in at.error[0].value  # nome não é exibido; só a linha/motivo
    assert "Linha 1" in at.error[0].value


def test_modelo_interno_ausente_mostra_mensagem_clara(monkeypatch, tmp_path):
    import src.generator as gerador

    monkeypatch.setattr(gerador, "DEFAULT_TEMPLATE_PATH", tmp_path / "sumiu.xlsx")
    at = novo_app()
    assert not at.exception
    assert "modelo interno não foi encontrado" in at.error[0].value
    assert len(at.button) == 0 and len(at.text_area) == 0

"""Atualizador de Promoções — cole o resultado da consulta e baixe a planilha."""

from __future__ import annotations

import base64
import json
import logging
import re
import uuid

import streamlit as st

from src.excel_service import ExcelServiceError, TemplateError
from src.generator import InvalidInput, generate_workbook, load_template

logger = logging.getLogger(__name__)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="Atualizador de Promoções", page_icon="📋", layout="centered")


def md(text: str) -> str:
    """Escapa caracteres especiais de markdown (ex.: '$', '_', '*') em textos dos dados."""
    return re.sub(r"([\\`*_{}\[\]()#+\-.!$~<>|:])", r"\\\1", text)


def start_download(filename: str, data: bytes) -> None:
    """Inicia o download no navegador e deixa um link discreto caso ele seja bloqueado."""
    payload = base64.b64encode(data).decode("ascii")
    page = f"""
<style>
  body {{ margin: 0; font-family: "Source Sans Pro", sans-serif; font-size: 14px; color: #808495; }}
  a {{ color: inherit; }}
</style>
<a id="dl" download={json.dumps(filename)} href="#">Se o download não iniciar, clique aqui.</a>
<script>
  // {uuid.uuid4().hex}  (garante que cada geração inicie um novo download)
  const bytes = Uint8Array.from(atob("{payload}"), c => c.charCodeAt(0));
  const url = URL.createObjectURL(new Blob([bytes], {{ type: {json.dumps(XLSX_MIME)} }}));
  const link = document.getElementById("dl");
  link.href = url;
  link.click();
</script>
"""
    st.iframe(page.strip(), height=28)


st.markdown("#### Atualizador de Promoções")

try:
    load_template()  # falha cedo, com mensagem clara, se o modelo interno sumiu
except TemplateError as exc:
    st.error(md(str(exc)))
    st.stop()

text = st.text_area(
    "Resultado da consulta SQL Server",
    height=400,
    label_visibility="collapsed",
    placeholder="Cole aqui o resultado da consulta SQL Server",
)

if st.button("Gerar e baixar planilha", type="primary"):
    try:
        result = generate_workbook(text)
    except InvalidInput as exc:
        itens = "\n".join(f"- {md(linha)}" for linha in exc.lines)
        extra = f"\n- … e mais {exc.total - len(exc.lines)} erro(s)." if exc.total > len(exc.lines) else ""
        st.error(f"Não foi possível gerar a planilha:\n\n{itens}{extra}")
    except ExcelServiceError as exc:
        st.error(md(str(exc)))
    except Exception as exc:  # noqa: BLE001
        logger.error("Falha inesperada na geração: %s", type(exc).__name__)
        st.error("Ocorreu um erro inesperado ao gerar a planilha.")
    else:
        st.success(f"Planilha gerada com {result.record_count} registro(s): {md(result.filename)}")
        if result.warnings:
            st.warning("Avisos (não impediram a geração):\n\n" + "\n".join(f"- {md(w)}" for w in result.warnings))
        start_download(result.filename, result.data)

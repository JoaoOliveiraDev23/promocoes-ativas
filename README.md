# Atualizador de Promoções

Cole o resultado da consulta SQL Server, clique em **Gerar e baixar planilha** e pronto:
o navegador baixa `PROMOCOES-ATIVAS DD-MM-AAAA.xlsx` (data de hoje) com os registros novos
no lugar dos antigos.

* A página tem apenas um campo de texto e um botão. Não há upload, validação manual nem confirmações.
* A validação acontece sozinha ao clicar. Se houver erro, a mensagem mostra a **linha** e o **motivo**.
* O modelo da planilha fica dentro do projeto, em `assets/PROMOCOES-ATIVAS-modelo.xlsx`, e é
  apenas **lido**: nunca é alterado ou sobrescrito. A saída é gerada em memória e baixada.
* Tudo roda no seu computador (sem banco de dados, login, nuvem ou APIs externas).

## Instalação (uma vez)

Pré-requisitos: Windows, [Python 3.11+](https://www.python.org/downloads/windows/)
(marque **"Add python.exe to PATH"** na instalação) e, opcionalmente, o VS Code.

1. Extraia a pasta `atualizador-promocoes` e abra-a no VS Code (**Arquivo → Abrir Pasta…**).
2. Abra o terminal (**Terminal → Novo Terminal**) e execute:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install -r requirements.txt
   ```

   Se o PowerShell bloquear a ativação ("execução de scripts desabilitada"), não altere
   proteções do sistema. Use uma destas alternativas seguras:

   * **Prompt de Comando:** na seta ao lado do `+` do terminal escolha *Command Prompt* e use
     `.venv\Scripts\activate.bat`.
   * **Sem ativar:** troque `python` por `.\.venv\Scripts\python.exe` em todos os comandos
     (ex.: `.\.venv\Scripts\python.exe -m streamlit run app.py`).

3. (Recomendado) Rode os testes: `python -m pytest -q`

## Uso

```powershell
python -m streamlit run app.py
```

O navegador abre em **http://localhost:8501** (se não abrir, copie o endereço). Na primeira
execução o Streamlit pode pedir um e-mail no terminal: pressione **Enter**.

1. No SQL Server Management Studio, execute a consulta, selecione o resultado na grade e copie
   com **Ctrl+C**, **sem cabeçalhos**. As datas devem estar como `AAAA-MM-DD`.
2. Cole no campo de texto.
3. Clique em **Gerar e baixar planilha**.

O download começa automaticamente. Se o navegador bloqueá-lo, aparece abaixo da mensagem de
sucesso o link "Se o download não iniciar, clique aqui". Para encerrar a aplicação, pressione
**Ctrl+C** no terminal.

Para treinar sem risco, cole o conteúdo de `exemplo_consulta.txt` (dados fictícios).

## Regras aplicadas automaticamente

* Seis colunas separadas por tabulação, nesta ordem: `COD`, `NOME`, `DETALHES`, `DATA INICIAL`,
  `DATA FINAL`, `PRIORIDADE`.
* Os registros colados **substituem integralmente** os do modelo; nenhum é filtrado ou descartado.
* `COD` e `PRIORIDADE` são gravados como número (inteiros, sem zeros à esquerda); as datas, como
  datas do Excel, e devem vir em `AAAA-MM-DD`; `NOME` e `DETALHES` são texto exatamente como colados.
  Só `DETALHES` pode ficar vazio.
* Texto que comece com `=`, `+`, `-` ou `@` é gravado como texto, nunca como fórmula.
* Cabeçalhos colados por engano, linhas com número errado de colunas, datas inválidas e linhas em
  branco no meio dos dados **impedem** a geração e são apontados pela linha.
* Datas finais anteriores às iniciais e `COD` repetido geram apenas um aviso (não impedem a geração).
* A tabela estruturada `Tabela1` é ajustada ao número de registros; cabeçalhos, formatação,
  larguras das colunas e nome da aba são preservados. Antes de entregar, o arquivo gerado é
  reaberto e conferido; se algo divergir, nada é baixado.
* A planilha gerada abre posicionada no topo (A1).

**Limitação do openpyxl:** o arquivo binário de configuração da impressora
(`printerSettings1.bin`) não é preservado; tamanho de papel e orientação são mantidos.

## Trocar o modelo

Substitua `assets/PROMOCOES-ATIVAS-modelo.xlsx` por outro arquivo **com o mesmo nome**, mantendo
os cabeçalhos `COD, NOME, DETALHES, DATA INICIAL, DATA FINAL, PRIORIDADE` na linha 1. Se o arquivo
estiver ausente ou inválido, a aplicação mostra uma mensagem clara.

> O modelo atual contém os registros antigos de 05/10/2026. Eles são sempre substituídos na
> saída, mas lembre-se disso se for compartilhar a pasta do projeto.

## Problemas comuns

| Sintoma | Solução |
|---|---|
| `python` não é reconhecido | Reinstale o Python marcando "Add python.exe to PATH", ou use `py`. |
| `No module named streamlit` | Ative o `.venv` e rode `python -m pip install -r requirements.txt`. |
| "O arquivo de modelo interno não foi encontrado" | Restaure `assets/PROMOCOES-ATIVAS-modelo.xlsx` na pasta do projeto. |
| "esperadas 6 colunas, encontradas N" | Texto sem tabulações, coluna a mais/menos ou quebra de linha dentro de um campo. |
| "parece a linha de cabeçalhos" | Apague a primeira linha ou copie sem cabeçalhos. |
| "formato dia/mês/ano não aceito" / "data com horário" | Na consulta, converta com `CONVERT(varchar(10), campo, 23)`. |
| Download não inicia | Use o link abaixo da mensagem de sucesso e permita downloads para `localhost`. |
| Porta 8501 em uso | `python -m streamlit run app.py --server.port 8502` |

## Estrutura

```
atualizador-promocoes/
├── app.py                                # interface: um campo e um botão
├── assets/PROMOCOES-ATIVAS-modelo.xlsx   # modelo interno (somente leitura)
├── requirements.txt
├── exemplo_consulta.txt
├── .streamlit/config.toml
├── src/
│   ├── generator.py     # texto -> validação -> planilha, em uma chamada
│   ├── data_parser.py   # leitura do texto colado (TSV)
│   ├── validators.py    # validação/conversão dos campos
│   ├── excel_service.py # atualização do modelo, conferência, nome do arquivo
│   ├── constants.py
│   └── models.py
└── tests/               # pytest
```

## Testes

`python -m pytest -q` executa os testes automatizados (parser, validadores, serviço Excel,
fluxo completo com o modelo interno e interface). Os testes usam dados em memória e arquivos
temporários e não alteram o modelo. `tests/test_browser.py` é opcional: usa um navegador real
(Playwright) e é ignorado se ele não estiver instalado.

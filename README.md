# Atualizador de Promoções (React + TypeScript)

Cole o resultado da consulta SQL Server, clique em **Gerar e baixar planilha** e o navegador baixa
`PROMOCOES-ATIVAS DD-MM-AAAA.xlsx` (data de hoje) com os registros novos no lugar dos antigos.

Mesma lógica da versão Python/Streamlit, agora 100% no navegador: **não há servidor, banco de dados,
login nem Python**. O site é só arquivos estáticos (ideal para um *Static Site* do Render), e os
dados colados **nunca saem do computador do usuário**.

## Requisitos

* [Node.js](https://nodejs.org/) 20 ou superior (22 recomendado) — inclui o `npm`.

## Rodar localmente

```bash
npm install        # uma vez
npm run dev        # desenvolvimento: http://localhost:5173 (recarrega ao salvar)
```

Outros comandos:

| Comando | O que faz |
|---|---|
| `npm test` | Roda os testes automatizados (Vitest) |
| `npm run typecheck` | Confere os tipos do TypeScript |
| `npm run build` | Confere tipos e gera o site de produção em `dist/` |
| `npm run preview` | Serve o `dist/` localmente (simula o que será publicado) |

## Como usar

1. No SQL Server Management Studio, execute a consulta, selecione o resultado na grade e copie com
   **Ctrl+C**, **sem cabeçalhos**. As datas devem estar como `AAAA-MM-DD`.
2. Cole no campo de texto.
3. Clique em **Gerar e baixar planilha**.

O download começa sozinho. Se o navegador o bloquear, aparece o link "Se o download não iniciar,
clique aqui". Em caso de erro, a mensagem mostra a **linha** e o **motivo** (até 5 erros por vez).
Para treinar, cole o conteúdo de `exemplo_consulta.txt` (dados fictícios).

## Publicar no Render

1. Envie o projeto para um repositório no GitHub/GitLab (a pasta `node_modules` e `dist` já são ignoradas).
2. No Render: **New → Static Site** e conecte o repositório, com:
   * **Build Command:** `npm ci && npm run build`
   * **Publish Directory:** `dist`
3. Pronto: cada `git push` publica uma nova versão automaticamente.

O arquivo `render.yaml` traz a mesma configuração (mais alguns cabeçalhos de segurança) para quem
preferir usar **Blueprints**. O `.node-version` indica o Node 22.

> **Atenção — o site fica público.** Qualquer pessoa com o endereço consegue abrir a página e usar a
> ferramenta. Ela não guarda nem recebe dados (tudo ocorre no navegador de quem usa), mas o modelo
> `public/PROMOCOES-ATIVAS-modelo.xlsx` é baixável por qualquer um que saiba o endereço. Por isso o
> modelo incluído **não contém os registros antigos** (eles são sempre substituídos mesmo). Os
> metadados do arquivo (aba *Propriedades* do Excel) ainda trazem o nome da empresa do autor
> original; remova-os no Excel se isso for um problema. Se precisar restringir o acesso à
> página, use um serviço de controle de acesso na frente do site.

## Regras aplicadas automaticamente

* Seis colunas separadas por tabulação: `COD`, `NOME`, `DETALHES`, `DATA INICIAL`, `DATA FINAL`, `PRIORIDADE`.
* Os registros colados **substituem integralmente** os do modelo; nenhum é filtrado ou descartado.
* `COD` e `PRIORIDADE`: números inteiros, sem zeros à esquerda. Datas: `AAAA-MM-DD`, gravadas como
  datas reais do Excel. `NOME` e `DETALHES`: texto exatamente como colado. Só `DETALHES` pode ficar vazio
  (o texto `NULL` do SSMS vira vazio, com aviso).
* Texto iniciado por `=`, `+`, `-` ou `@` é gravado como texto, nunca como fórmula.
* Cabeçalhos colados por engano, número errado de colunas, datas inválidas e linhas em branco no meio
  **impedem** a geração. Data final anterior à inicial e `COD` repetido geram só **aviso**.
* A tabela estruturada `Tabela1` é ajustada ao número de registros; cabeçalhos, formatação, larguras e
  nome da aba são preservados. Antes de entregar, o arquivo gerado é relido e conferido; se algo
  divergir, nada é baixado.
* A planilha gerada abre posicionada no topo (A1).

### Como a planilha é editada

Em vez de recriar o arquivo, o programa abre o `.xlsx` (que é um ZIP de XMLs) e altera **somente**
a aba de dados (linhas 2 em diante), o intervalo da tabela e, se algum texto começar com `=`, `+`, `-`
ou `@`, acrescenta um estilo. Todo o restante do pacote (tema, estilos, larguras, configuração de
impressão etc.) é copiado **byte a byte**. A biblioteca `fflate` só cuida do ZIP.

## Trocar o modelo

Substitua `public/PROMOCOES-ATIVAS-modelo.xlsx` (mesmo nome) por outro arquivo com os cabeçalhos
`COD, NOME, DETALHES, DATA INICIAL, DATA FINAL, PRIORIDADE` na linha 1 e a linha 2 formatada como
quer que os dados saiam (as linhas novas copiam o estilo da linha 2). Se o modelo estiver ausente ou
inválido, a página mostra uma mensagem clara. Antes de publicar um modelo novo, apague os registros
antigos dele se não quiser expô-los.

## Personalizar o visual

| Onde | O quê |
|---|---|
| `src/styles.css` | **Comece pelas variáveis em `:root`**: cores, fonte, raio dos cantos, largura da página. O resto do arquivo está organizado por componente (`.title`, `.input`, `.button`, `.message`). |
| `src/App.tsx` | Estrutura da página, textos, estados e mensagens. |
| `src/components/Message.tsx` | Caixa de sucesso/erro/aviso. |
| `index.html` | Título da aba, ícone e idioma. |

O projeto não usa biblioteca de componentes. Se quiser, o PrimeReact pode ser adicionado depois
(`npm i primereact` e importar o tema em `main.tsx`), trocando `<textarea>`, `<button>` e `Message`
pelos equivalentes dele; a lógica em `src/core` não muda.

## Estrutura

```
├── index.html
├── public/PROMOCOES-ATIVAS-modelo.xlsx    # modelo interno (somente leitura)
├── exemplo_consulta.txt
├── render.yaml  .node-version
├── src/
│   ├── main.tsx  App.tsx  styles.css      # interface
│   ├── components/Message.tsx
│   ├── download.ts  loadTemplate.ts       # download e carregamento do modelo (navegador)
│   └── core/                              # regras (sem dependência do navegador)
│       ├── constants.ts  models.ts
│       ├── dataParser.ts                  # leitura do texto colado (TSV)
│       ├── validators.ts                  # validação/conversão dos campos
│       ├── generator.ts                   # texto -> validação -> planilha
│       ├── filename.ts                    # PROMOCOES-ATIVAS DD-MM-AAAA.xlsx
│       └── xlsx/                          # análise e edição do .xlsx
└── tests/                                 # Vitest
```

## Problemas comuns

| Sintoma | Solução |
|---|---|
| "O arquivo de modelo interno não foi encontrado" | Restaure `public/PROMOCOES-ATIVAS-modelo.xlsx` e publique novamente. |
| "esperadas 6 colunas, encontradas N" | Texto sem tabulações, coluna a mais/menos ou quebra de linha dentro de um campo. |
| "parece a linha de cabeçalhos" | Apague a primeira linha ou copie sem cabeçalhos. |
| "formato dia/mês/ano não aceito" / "data com horário" | Na consulta, converta com `CONVERT(varchar(10), campo, 23)`. |
| "zeros à esquerda" | A coluna é numérica no modelo; corrija o dado de origem. |
| Download não inicia | Use o link abaixo da mensagem de sucesso e permita downloads para o site. |
| Build no Render falha por versão do Node | Em *Environment*, defina `NODE_VERSION` = `22`. |

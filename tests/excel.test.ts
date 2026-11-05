import { readFileSync } from "node:fs";
import { strFromU8, strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { parsePromotions } from "../src/core/dataParser";
import { excelSerial, GenerationError, TemplateError, updateWorkbook } from "../src/core/xlsx";
import { buildModel, HEADERS } from "./helpers/buildModel";
import { makeRecords } from "./helpers/records";
import { readOutput } from "./helpers/readOutput";

const SERIAL_2026_01 = excelSerial({ year: 2026, month: 1, day: 1 });

describe("excelSerial", () => {
  it("bate com os números de série reais do modelo", () => {
    expect(excelSerial({ year: 2025, month: 3, day: 10 })).toBe(45726);
    expect(excelSerial({ year: 2026, month: 12, day: 31 })).toBe(46387);
  });
  it("trata o 'bug' do ano bissexto de 1900 do Excel", () => {
    expect(excelSerial({ year: 1900, month: 1, day: 1 })).toBe(1);
    expect(excelSerial({ year: 1900, month: 2, day: 28 })).toBe(59);
    expect(excelSerial({ year: 1900, month: 3, day: 1 })).toBe(61);
  });
});

describe("substituição dos dados", () => {
  it("substitui integralmente os registros antigos", () => {
    const out = readOutput(updateWorkbook(buildModel({ records: 5, tableRows: 20 }), makeRecords(5)));
    const novos = makeRecords(5);
    novos.forEach((rec, i) => {
      expect(out.row(2 + i)).toEqual([
        rec.cod,
        rec.nome,
        rec.detalhes,
        excelSerial(rec.dataInicial),
        excelSerial(rec.dataFinal),
        rec.prioridade,
      ]);
    });
    expect(out.sheetXml).not.toContain("Antiga");
    expect(out.sheetXml).not.toContain("9000");
  });

  it("remove registros excedentes antigos e encolhe a tabela", () => {
    const out = readOutput(updateWorkbook(buildModel({ records: 15, tableRows: 15 }), makeRecords(3)));
    expect(out.row(4)[0]).toBe(1002);
    expect(out.rowsWithData()).toEqual([2, 3, 4]);
    expect(out.rowsWithAnyCell()).toEqual([2, 3, 4]); // nem linhas vazias formatadas sobram
    expect(out.tableXml).toContain('ref="A1:F4"');
  });

  it("inclui novas linhas com a formatação padrão do modelo", () => {
    const out = readOutput(updateWorkbook(buildModel({ records: 5, tableRows: 20 }), makeRecords(40)));
    expect(out.row(41)[0]).toBe(1039);
    expect(out.row(42)).toEqual([null, null, null, null, null, null]);
    expect(out.tableXml).toContain('ref="A1:F41"');
    const padrao = out.styles(2);
    for (const r of [3, 21, 22, 30, 41]) expect(out.styles(r)).toEqual(padrao);
    expect(padrao).toEqual(["3", "2", "2", "4", "4", "3"]); // centralizado / esquerda / data / ...
  });

  it("COD e PRIORIDADE são números, datas são séries, textos são texto", () => {
    const out = readOutput(updateWorkbook(buildModel(), makeRecords(1)));
    const [cod, nome, det, ini, fim, prio] = out.row(2);
    expect(typeof cod).toBe("number");
    expect(typeof nome).toBe("string");
    expect(typeof det).toBe("string");
    expect(ini).toBe(SERIAL_2026_01);
    expect(fim).toBe(excelSerial({ year: 2026, month: 12, day: 31 }));
    expect(typeof prio).toBe("number");
    expect(out.styles(2)[3]).toBe("4"); // estilo com formato de data do modelo
  });

  it("DETALHES vazio vira célula vazia, mantendo a formatação", () => {
    const rec = { ...makeRecords(1)[0], detalhes: "" };
    const out = readOutput(updateWorkbook(buildModel(), [rec]));
    expect(out.row(2)[2]).toBeNull();
    expect(out.styles(2)[2]).toBe("2");
  });

  it("preserva CODs grandes sem perda de precisão", () => {
    const rec = { ...makeRecords(1)[0], cod: 999999999999999 };
    expect(readOutput(updateWorkbook(buildModel(), [rec])).row(2)[0]).toBe(999999999999999);
  });
});

describe("segurança do conteúdo", () => {
  it("texto que começa com = + - @ nunca vira fórmula e ganha quotePrefix", () => {
    const parsed = parsePromotions('1\t=1+1\t=HYPERLINK("http://x","y")\t2026-01-01\t2026-01-02\t1\n2\t-30% Natal\t+55 11\t2026-01-01\t2026-01-02\t1\n3\t@menção\tnormal\t2026-01-01\t2026-01-02\t1');
    const bytes = updateWorkbook(buildModel(), parsed.records);
    const out = readOutput(bytes);
    expect(out.row(2)[1]).toBe("=1+1");
    expect(out.row(2)[2]).toBe('=HYPERLINK("http://x","y")');
    expect(out.row(3)[1]).toBe("-30% Natal");
    expect(out.sheetXml).not.toMatch(/<f[\s>/]/);
    // estilo-variante com quotePrefix criado a partir do estilo da coluna (2), contagem atualizada
    expect(out.stylesXml).toMatch(/<cellXfs count="6">/);
    expect(out.stylesXml).toContain('quotePrefix="1"');
    expect(out.styles(2)[1]).toBe("5");
    expect(out.styles(4)[2]).toBe("2"); // "normal" não precisa
  });

  it("sem textos especiais, styles.xml permanece idêntico", () => {
    const model = buildModel();
    const out = readOutput(updateWorkbook(model, makeRecords(3)));
    expect(out.stylesXml).toBe(strFromU8(readOutput(model).files["xl/styles.xml"]));
  });

  it("escapa XML e preserva sequências _xHHHH_ e quebras de linha", () => {
    const rec = { ...makeRecords(1)[0], nome: "A & B <c> \"d\" 'e' _x000D_ _x0041_", detalhes: "linha 1\nlinha 2\tcom tab  e  espaços " };
    const out = readOutput(updateWorkbook(buildModel(), [rec]));
    expect(out.row(2)[1]).toBe(rec.nome);
    expect(out.row(2)[2]).toBe(rec.detalhes);
    expect(out.sheetXml).toContain("_x005F_x000D_"); // escapado para o Excel não reinterpretar
  });

  it("emoji e caracteres fora do BMP são preservados", () => {
    const rec = { ...makeRecords(1)[0], nome: "Oferta 😀 𝒳 ção" };
    expect(readOutput(updateWorkbook(buildModel(), [rec])).row(2)[1]).toBe(rec.nome);
  });
});

describe("preservação do modelo", () => {
  it("só altera aba, tabela (e estilos se preciso); demais partes ficam idênticas byte a byte", () => {
    const model = buildModel({ records: 5, tableRows: 20 });
    const antes = readOutput(model).files;
    const depois = readOutput(updateWorkbook(model, makeRecords(8))).files;
    expect(Object.keys(depois)).toEqual(Object.keys(antes));
    const iguais = (a: Uint8Array, b: Uint8Array) => a.length === b.length && a.every((v, i) => v === b[i]);
    const alteradas = Object.keys(antes).filter((n) => !iguais(antes[n], depois[n]));
    expect(alteradas.sort()).toEqual(["xl/tables/table1.xml", "xl/worksheets/sheet1.xml"]);
    expect([...depois["xl/printerSettings/printerSettings1.bin"]]).toEqual([...antes["xl/printerSettings/printerSettings1.bin"]]);
  });

  it("mantém cabeçalhos, larguras das colunas e demais elementos da aba", () => {
    const model = buildModel();
    const antes = readOutput(model).sheetXml;
    const depois = readOutput(updateWorkbook(model, makeRecords(8))).sheetXml;
    const pega = (xml: string, re: RegExp) => re.exec(xml)?.[0];
    expect(pega(depois, /<row r="1".*?<\/row>/)).toBe(pega(antes, /<row r="1".*?<\/row>/));
    expect(pega(depois, /<cols>.*?<\/cols>/)).toBe(pega(antes, /<cols>.*?<\/cols>/));
    expect(pega(depois, /<sheetFormatPr[^>]*>/)).toBe(pega(antes, /<sheetFormatPr[^>]*>/));
    expect(pega(depois, /<pageSetup[^>]*>/)).toBe(pega(antes, /<pageSetup[^>]*>/));
    expect(pega(depois, /<pageMargins[^>]*>/)).toBe(pega(antes, /<pageMargins[^>]*>/));
    expect(depois.startsWith(antes.slice(0, antes.indexOf("<dimension")))).toBe(true);
  });

  it("tabela estruturada: atualiza ref e autoFilter e preserva nome, estilo e colunas", () => {
    const out = readOutput(updateWorkbook(buildModel(), makeRecords(9)));
    expect(out.tableXml).toContain('name="Tabela1"');
    expect(out.tableXml).toContain('displayName="Tabela1"');
    expect(out.tableXml).toMatch(/<table [^>]*ref="A1:F10"/);
    expect(out.tableXml).toContain('<autoFilter ref="A1:F10"');
    expect(out.tableXml).toContain('name="TableStyleMedium16"');
    expect(out.tableXml).toContain('showRowStripes="1"');
    for (const h of HEADERS) expect(out.tableXml).toContain(`name="${h}"`);
  });

  it("sortState da tabela acompanha o novo intervalo", () => {
    const out = readOutput(updateWorkbook(buildModel({ sortState: true, tableRows: 20 }), makeRecords(6)));
    expect(out.tableXml).toContain('<sortState ref="A2:F7">');
    expect(out.tableXml).toContain('<sortCondition ref="A2:A7"');
  });

  it("modelo sem tabela, com autofiltro de planilha: atualiza o filtro", () => {
    const out = readOutput(updateWorkbook(buildModel({ withTable: false, sheetAutoFilter: true, tableRows: 10 }), makeRecords(4)));
    expect(out.sheetXml).toContain('<autoFilter ref="A1:F5"/>');
  });

  it("modelo sem tabela funciona", () => {
    const out = readOutput(updateWorkbook(buildModel({ withTable: false, records: 5, tableRows: 5 }), makeRecords(7)));
    expect(out.row(8)[0]).toBe(1006);
    expect(out.tableXml).toBeUndefined();
  });

  it("janela volta ao topo (A1), sem painéis congelados", () => {
    const out = readOutput(updateWorkbook(buildModel(), makeRecords(3)));
    expect(out.sheetXml).not.toContain("topLeftCell");
    expect(out.sheetXml).toContain('<selection activeCell="A1" sqref="A1"/>');
  });

  it("painéis congelados são preservados", () => {
    const model = buildModel({ freeze: true });
    const antes = /<sheetViews>.*?<\/sheetViews>/.exec(readOutput(model).sheetXml)![0];
    const depois = /<sheetViews>.*?<\/sheetViews>/.exec(readOutput(updateWorkbook(model, makeRecords(3))).sheetXml)![0];
    expect(depois).toBe(antes);
  });

  it("dimensão acompanha os dados", () => {
    expect(readOutput(updateWorkbook(buildModel({ tableRows: 30 }), makeRecords(3))).sheetXml).toContain('<dimension ref="A1:F4"/>');
  });

  it("conteúdo fora de A:F é preservado e entra na dimensão", () => {
    const out = readOutput(updateWorkbook(buildModel({ extraCell: true }), makeRecords(2)));
    expect(out.sheetXml).toContain('<c r="G5" t="inlineStr"><is><t>observação</t></is></c>');
    expect(out.sheetXml).toMatch(/<row r="5"[^>]*>.*<c r="G5"/);
    expect(out.sheetXml).toContain('<dimension ref="A1:G5"/>');
    expect(out.row(5)).toEqual([null, null, null, null, null, null]); // dados antigos não voltam
  });

  it("célula vazia estilizada fora de A:F (G32) é mantida mesmo com poucos registros", () => {
    const out = readOutput(updateWorkbook(buildModel({ extraEmptyStyledG32: true, tableRows: 40 }), makeRecords(3)));
    expect(out.sheetXml).toContain('<c r="G32" s="3"/>');
  });

  it("outras abas são preservadas", () => {
    const model = buildModel();
    const files = readOutput(model).files;
    files["xl/worksheets/sheet2.xml"] = strToU8("<worksheet/>");
    files["xl/workbook.xml"] = strToU8(
      strFromU8(files["xl/workbook.xml"]).replace("</sheets>", '<sheet name="Notas" sheetId="2" r:id="rId9"/></sheets>'),
    );
    files["xl/_rels/workbook.xml.rels"] = strToU8(
      strFromU8(files["xl/_rels/workbook.xml.rels"]).replace(
        "</Relationships>",
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
      ),
    );
    const out = readOutput(updateWorkbook(zipSync(files), makeRecords(2)));
    expect(strFromU8(out.files["xl/worksheets/sheet2.xml"])).toBe("<worksheet/>");
  });

  it("não modifica os bytes recebidos", () => {
    const model = buildModel();
    const copia = new Uint8Array(model);
    updateWorkbook(model, makeRecords(3));
    expect([...model]).toEqual([...copia]);
  });

  it("gera o mesmo resultado com o mesmo modelo e os mesmos dados", () => {
    const model = buildModel();
    const a = readOutput(updateWorkbook(model, makeRecords(4)));
    const b = readOutput(updateWorkbook(model, makeRecords(4)));
    expect(a.sheetXml).toBe(b.sheetXml);
    expect(a.tableXml).toBe(b.tableXml);
  });
});

describe("modelos inválidos", () => {
  const gera = (m: Uint8Array) => () => updateWorkbook(m, makeRecords(2));

  it.each([
    ["vazio", new Uint8Array(), /vazio/],
    ["texto qualquer", strToU8("isto nao e um xlsx"), /não é um \.xlsx válido/],
    ["zip quebrado", new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4]), /não é um \.xlsx válido/],
  ])("rejeita arquivo %s", (_nome, bytes, msg) => {
    expect(gera(bytes)).toThrow(TemplateError);
    expect(gera(bytes)).toThrow(msg);
  });

  it("rejeita zip que não é Excel", () => {
    expect(gera(zipSync({ "qualquer.txt": strToU8("x") }))).toThrow(/estrutura/);
  });

  it("rejeita arquivo com macros (.xlsm)", () => {
    expect(gera(buildModel({ macro: true }))).toThrow(/macros/);
  });

  it("rejeita cabeçalhos inesperados e mostra os encontrados", () => {
    const m = buildModel({ headers: ["CODIGO", "NOME", "DETALHES", "INICIO", "FIM", "PRIORIDADE"] });
    expect(gera(m)).toThrow(/cabeçalhos esperados/);
    expect(gera(m)).toThrow(/CODIGO/);
  });

  it("rejeita fórmula na área de dados", () => {
    expect(gera(buildModel({ formulaInData: true }))).toThrow(/A2 contém fórmula/);
  });

  it("rejeita células mescladas na área de dados", () => {
    expect(gera(buildModel({ mergeInData: true }))).toThrow(/mescladas/);
  });

  it("rejeita outra tabela sobre as colunas A:F", () => {
    expect(gera(buildModel({ otherTableOverlapping: true }))).toThrow(/outra tabela/);
  });

  it("rejeita tabela com linha de totais", () => {
    expect(gera(buildModel({ totalsRow: true }))).toThrow(/linha de totais/);
  });

  it("rejeita modelo cujas colunas de data não têm formato de data", () => {
    expect(gera(buildModel({ dateStyle: "general" }))).toThrow(/formatadas como data/);
  });

  it("rejeita sistema de datas 1904", () => {
    expect(gera(buildModel({ date1904: true }))).toThrow(/1904/);
  });

  it("rejeita mais de uma aba com os mesmos cabeçalhos", () => {
    expect(gera(buildModel({ secondSheetWithSameHeaders: true }))).toThrow(/Mais de uma aba/);
  });

  it("sem registros não gera", () => {
    expect(() => updateWorkbook(buildModel(), [])).toThrow(GenerationError);
    expect(() => updateWorkbook(buildModel(), [])).toThrow(/Não há registros/);
  });
});

describe("modelo interno real (public/)", () => {
  const modelo = new Uint8Array(readFileSync("public/PROMOCOES-ATIVAS-modelo.xlsx"));

  it("é válido: aba Planilha1, Tabela1 e nenhum registro antigo", () => {
    const out = readOutput(modelo);
    expect(out.tableXml).toContain('displayName="Tabela1"');
    expect(out.rowsWithData()).toEqual([]);
    expect(out.row(1)).toEqual(HEADERS);
  });

  it("não contém textos de promoções (somente os 6 cabeçalhos)", () => {
    expect(readOutput(modelo).sst).toEqual(HEADERS);
  });

  it("gera planilha a partir do modelo real e confere o resultado", () => {
    const out = readOutput(updateWorkbook(modelo, makeRecords(98)));
    expect(out.rowsWithData()).toHaveLength(98);
    expect(out.rowsWithAnyCell()).toHaveLength(98);
    expect(out.tableXml).toContain('ref="A1:F99"');
    expect(out.styles(2)).toEqual(["6", "9", "9", "5", "5", "6"]); // mesmos estilos do modelo original
    expect(out.styles(99)).toEqual(out.styles(2));
    expect(out.sheetXml).toContain('<dimension ref="A1:G99"/>');
    expect(out.sheetXml).toContain('<c r="G32" s="3"/>');
  });

  it("com mais registros que o modelo, expande com a formatação padrão", () => {
    const out = readOutput(updateWorkbook(modelo, makeRecords(300)));
    expect(out.rowsWithData()).toHaveLength(300);
    expect(out.styles(301)).toEqual(out.styles(2));
    expect(out.tableXml).toContain('ref="A1:F301"');
  });
});

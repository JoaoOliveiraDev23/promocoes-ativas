import { describe, expect, it } from "vitest";
import { parsePromotions, splitTsv } from "../src/core/dataParser";
import { isValid } from "../src/core/models";

const HEADER = "COD\tNOME\tDETALHES\tDATA INICIAL\tDATA FINAL\tPRIORIDADE";
const ROW1 = "187\tPetit Deli Ciclo 3 - 2025\tGranéis até 12 Unidades - R$ 2,87\t2025-03-10\t2026-12-31\t4";
const ROW2 = "233\tKIT VITALÍCIO VD 1ª COMPRA\tPrimeira compra\t2025-05-05\t2026-12-31\t1";

describe("parsePromotions", () => {
  it("importa registros válidos separados por tabulação", () => {
    const r = parsePromotions(`${ROW1}\n${ROW2}`);
    expect(isValid(r)).toBe(true);
    expect(r.errors).toEqual([]);
    expect(r.warnings).toEqual([]);
    expect(r.records).toHaveLength(2);
    expect(r.records[0].cod).toBe(187);
    expect(r.records[0].prioridade).toBe(4);
    expect(r.records[0].dataInicial).toEqual({ year: 2025, month: 3, day: 10 });
    expect(r.records[0].dataFinal).toEqual({ year: 2026, month: 12, day: 31 });
  });

  it("preserva acentos e caracteres especiais", () => {
    const t = "Ç ª á ã é í ú – “aspas” R$ 1.249,74 (50%) & <b> ;,";
    const r = parsePromotions(`1\tNome ${t}\t${t}\t2026-01-01\t2026-01-02\t1`);
    expect(isValid(r)).toBe(true);
    expect(r.records[0].detalhes).toBe(t);
    expect(r.records[0].nome).toBe(`Nome ${t}`);
  });

  it("preserva espaços internos e nas pontas do texto", () => {
    const r = parsePromotions("1\t  Nome  com  espaços \tDetalhe \t2026-01-01\t2026-01-02\t1");
    expect(r.records[0].nome).toBe("  Nome  com  espaços ");
    expect(r.records[0].detalhes).toBe("Detalhe ");
  });

  it("aceita campo vazio em DETALHES", () => {
    const r = parsePromotions("1\tNome\t\t2026-01-01\t2026-01-02\t1");
    expect(isValid(r)).toBe(true);
    expect(r.records[0].detalhes).toBe("");
  });

  it("campos obrigatórios vazios geram erro com linha e coluna", () => {
    const r = parsePromotions(`${ROW1}\n\t\tx\t\t2026-01-02\t`);
    expect(isValid(r)).toBe(false);
    const colunas = new Set(r.errors.filter((e) => e.line === 2).map((e) => e.column));
    expect(colunas).toEqual(new Set(["COD", "NOME", "DATA INICIAL", "PRIORIDADE"]));
  });

  it("quantidade incorreta de colunas", () => {
    const r = parsePromotions(`${ROW1}\n1\tNome\t2026-01-01\t2026-01-02\t1\n${ROW2}\n1\t2\t3\t4\t5\t6\t7`);
    expect(isValid(r)).toBe(false);
    expect(r.errors.map((e) => e.line)).toEqual([2, 4]);
    expect(r.errors[0].message).toContain("esperadas 6 colunas, encontradas 5");
    expect(r.errors[1].message).toContain("encontradas 7");
    expect(r.records).toHaveLength(2); // as linhas boas continuam identificadas
  });

  it("texto sem tabulação explica o problema", () => {
    const r = parsePromotions("187 Petit Deli 2025-03-10 2026-12-31 4");
    expect(isValid(r)).toBe(false);
    expect(r.errors[0].message).toContain("Nenhuma tabulação");
  });

  it("identifica cabeçalho colado sem corromper", () => {
    const r = parsePromotions(`${HEADER}\n${ROW1}`);
    expect(isValid(r)).toBe(false);
    expect(r.errors[0].line).toBe(1);
    expect(r.errors[0].message).toContain("cabeçalhos");
    expect(r.errors[0].message).toContain("sem cabeçalhos");
    expect(r.records).toHaveLength(1);
  });

  it("detecta cabeçalho fora de ordem ou em minúsculas", () => {
    const r = parsePromotions(`nome\tcod\tdetalhes\tprioridade\tdata final\tdata inicial\n${ROW1}`);
    expect(isValid(r)).toBe(false);
    expect(r.errors[0].message).toContain("cabeçalhos");
  });

  it("ignora linhas em branco no início e no fim", () => {
    const r = parsePromotions(`\n\n  \n${ROW1}\n${ROW2}\n\n\n   \n`);
    expect(isValid(r)).toBe(true);
    expect(r.records).toHaveLength(2);
    expect(r.rowsRead).toBe(2);
  });

  it("numeração de linhas considera as linhas em branco iniciais", () => {
    const r = parsePromotions(`\n\n${ROW1}\n1\tx\t2026-01-01\t1`);
    expect(r.errors.map((e) => e.line)).toEqual([4]);
  });

  it("linha em branco no meio é erro", () => {
    const r = parsePromotions(`${ROW1}\n\n${ROW2}`);
    expect(isValid(r)).toBe(false);
    expect(r.errors[0].line).toBe(2);
    expect(r.errors[0].message).toContain("em branco");
  });

  it.each(["", "\n\n", "   \t \n"])("entrada vazia (%j) é erro", (texto) => {
    const r = parsePromotions(texto);
    expect(isValid(r)).toBe(false);
    expect(r.errors[0].message.toLowerCase()).toContain("nenhum dado");
  });

  it.each(["\r\n", "\r"])("aceita quebras de linha %j", (sep) => {
    const r = parsePromotions(`${ROW1}${sep}${ROW2}${sep}`);
    expect(isValid(r)).toBe(true);
    expect(r.records).toHaveLength(2);
  });

  it("ignora BOM no início", () => {
    expect(isValid(parsePromotions("\ufeff" + ROW1))).toBe(true);
  });

  it("campo entre aspas com quebra de linha, como no Excel", () => {
    const texto = `1\tNome\t"linha 1\nlinha 2 com ""aspas"""\t2026-01-01\t2026-01-02\t1\n${ROW2}`;
    const r = parsePromotions(texto);
    expect(isValid(r)).toBe(true);
    expect(r.records[0].detalhes).toBe('linha 1\nlinha 2 com "aspas"');
    expect(r.records).toHaveLength(2);
  });

  it("numeração de linhas após campo multilinha", () => {
    const r = parsePromotions('1\tNome\t"a\nb\nc"\t2026-01-01\t2026-01-02\t1\n1\tx\t2026-01-01\t1');
    expect(r.errors.map((e) => e.line)).toEqual([4]);
  });

  it("aspa solta no meio do texto não altera o valor", () => {
    const r = parsePromotions('1\t"Natal" especial\tTem 5" de altura\t2026-01-01\t2026-01-02\t1');
    expect(isValid(r)).toBe(true);
    expect(r.records[0].nome).toBe('"Natal" especial');
    expect(r.records[0].detalhes).toBe('Tem 5" de altura');
  });

  it("NULL em DETALHES vira vazio com aviso; em obrigatório é erro", () => {
    const ok = parsePromotions("1\tNome\tNULL\t2026-01-01\t2026-01-02\t1");
    expect(isValid(ok)).toBe(true);
    expect(ok.records[0].detalhes).toBe("");
    expect(ok.warnings[0].severity).toBe("aviso");
    const ruim = parsePromotions("NULL\tNome\tx\t2026-01-01\t2026-01-02\t1");
    expect(isValid(ruim)).toBe(false);
    expect(ruim.errors[0].column).toBe("COD");
    expect(ruim.errors[0].message).toContain("NULL");
  });

  it("COD repetido e datas invertidas são apenas avisos", () => {
    const r = parsePromotions("1\tA\tx\t2026-05-01\t2026-04-01\t1\n1\tB\tx\t2026-01-01\t2026-02-01\t1");
    expect(isValid(r)).toBe(true);
    const msgs = r.warnings.map((w) => w.message).join(" | ");
    expect(msgs).toContain("anterior");
    expect(msgs).toContain("mais de uma vez");
  });

  it("não filtra registros por data: datas passadas e futuras são mantidas", () => {
    const r = parsePromotions("1\tA\tx\t2001-01-01\t2001-12-31\t1\n2\tB\tx\t2099-01-01\t2099-12-31\t1");
    expect(isValid(r)).toBe(true);
    expect(r.records).toHaveLength(2);
  });
});

describe("splitTsv", () => {
  it("não cria linha extra no fim", () => {
    expect(splitTsv("a\tb\nc\td\n")).toHaveLength(2);
    expect(splitTsv("a\tb\n\nc").map((r) => r.fields)).toEqual([["a", "b"], [""], ["c"]]);
  });
});

import { describe, expect, it } from "vitest";
import {
  FieldError,
  looksLikeHeader,
  parseInteger,
  parseIsoDate,
  validateRow,
} from "../src/core/validators";
import { describeIssue } from "../src/core/models";

const row = (line: number, ...fields: string[]) => ({ line, fields });
const GOOD = ["10", "Nome", "Detalhe", "2026-01-31", "2026-12-31", "2"];

describe("datas", () => {
  it("aceita datas válidas, inclusive bissexto", () => {
    expect(parseIsoDate("2026-02-28")).toEqual({ year: 2026, month: 2, day: 28 });
    expect(parseIsoDate(" 2024-02-29 ")).toEqual({ year: 2024, month: 2, day: 29 });
  });

  it.each([
    ["", "obrigatório"],
    ["2026-02-30", "inexistente"],
    ["2025-02-29", "inexistente"],
    ["2026-13-01", "inexistente"],
    ["10/03/2025", "dia/mês/ano"],
    ["2026/03/10", "inválida"],
    ["26-03-10", "inválida"],
    ["2026-3-1", "inválida"],
    ["2026-03-10 00:00:00.000", "horário"],
    ["2026-03-10T10:00", "horário"],
    ["1899-12-31", "anterior"],
    ["abc", "inválida"],
  ])("rejeita %j (%s)", (valor, trecho) => {
    expect(() => parseIsoDate(valor)).toThrow(FieldError);
    expect(() => parseIsoDate(valor)).toThrow(new RegExp(trecho));
  });
});

describe("inteiros", () => {
  it("aceita inteiros válidos", () => {
    expect(parseInteger("187")).toBe(187);
    expect(parseInteger(" 4 ")).toBe(4);
    expect(parseInteger("0")).toBe(0);
    expect(parseInteger("999999999999999")).toBe(999999999999999);
  });

  it.each([
    ["", "obrigatório"],
    ["1,5", "inteiro"],
    ["1.0", "inteiro"],
    ["-1", "inteiro"],
    ["abc", "inteiro"],
    ["1 2", "inteiro"],
    ["0123", "zeros à esquerda"],
    ["1".repeat(16), "grande demais"],
    ["١٢٣", "inteiro"], // dígitos arábico-índicos
  ])("rejeita %j (%s)", (valor, trecho) => {
    expect(() => parseInteger(valor)).toThrow(FieldError);
    expect(() => parseInteger(valor)).toThrow(new RegExp(trecho));
  });
});

describe("cabeçalho", () => {
  it("detecta cabeçalho exato", () => {
    expect(looksLikeHeader(["COD", "NOME", "DETALHES", "DATA INICIAL", "DATA FINAL", "PRIORIDADE"])).toBe(true);
  });
  it("ignora caixa e espaços", () => {
    expect(looksLikeHeader([" cod ", "Nome", "detalhes", "data  inicial", "x", "y"])).toBe(true);
  });
  it("registro normal não é cabeçalho", () => {
    expect(looksLikeHeader(GOOD)).toBe(false);
  });
  it("promoção chamada NOME não é cabeçalho", () => {
    expect(looksLikeHeader(["5", "NOME", "COD", "2026-01-01", "2026-01-02", "1"])).toBe(false);
  });
});

describe("validateRow", () => {
  it("linha válida", () => {
    const { record, errors, warnings } = validateRow(row(7, ...GOOD));
    expect(errors).toEqual([]);
    expect(warnings).toEqual([]);
    expect([record!.cod, record!.nome, record!.prioridade]).toEqual([10, "Nome", 2]);
  });

  it("erros informam linha e coluna", () => {
    const { record, errors } = validateRow(row(42, "x", "Nome", "d", "31/01/2026", "2026-12-31", "alta"));
    expect(record).toBeNull();
    expect(new Set(errors.map((e) => `${e.line}|${e.column}`))).toEqual(
      new Set(["42|COD", "42|DATA INICIAL", "42|PRIORIDADE"]),
    );
    expect(describeIssue(errors[0])).toContain("Linha 42 [COD]");
  });

  it("caractere de controle é rejeitado", () => {
    const { errors } = validateRow(row(1, "1", "No\u0000me", "d", "2026-01-01", "2026-01-02", "1"));
    expect(errors[0].column).toBe("NOME");
    expect(errors[0].message).toContain("controle");
  });

  it("substituto (surrogate) solto é rejeitado, par válido (emoji) é aceito", () => {
    expect(validateRow(row(1, "1", "ok \ud83d", "d", "2026-01-01", "2026-01-02", "1")).errors[0].column).toBe("NOME");
    expect(validateRow(row(1, "1", "ok 😀", "d", "2026-01-01", "2026-01-02", "1")).errors).toEqual([]);
  });

  it("texto longo demais", () => {
    const { errors } = validateRow(row(1, "1", "N", "x".repeat(32768), "2026-01-01", "2026-01-02", "1"));
    expect(errors[0].column).toBe("DETALHES");
    expect(errors[0].message).toContain("32767");
  });

  it("nome só com espaços é vazio", () => {
    const { errors } = validateRow(row(1, "1", "   ", "d", "2026-01-01", "2026-01-02", "1"));
    expect(errors[0].column).toBe("NOME");
  });

  it("fórmula digitada é aceita como texto", () => {
    const { record, errors } = validateRow(row(1, "1", "=1+1", '=HYPERLINK("x")', "2026-01-01", "2026-01-02", "1"));
    expect(errors).toEqual([]);
    expect(record!.nome).toBe("=1+1");
  });
});

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { buildOutputFilename } from "../src/core/filename";
import { generateWorkbook, InvalidInput } from "../src/core/generator";
import { ExcelServiceError, TemplateError } from "../src/core/xlsx";
import { readOutput } from "./helpers/readOutput";

const modelo = new Uint8Array(readFileSync("public/PROMOCOES-ATIVAS-modelo.xlsx"));
const LINHA = (cod: number, prio: number) => `${cod}\tPromoção ${cod} – R$ 2,87\tDetalhe ${cod}\t2026-03-10\t2026-12-31\t${prio}`;
const texto = (n: number) => Array.from({ length: n }, (_v, i) => LINHA(500 + i, 1 + (i % 4))).join("\n") + "\n";

describe("nome do arquivo", () => {
  it("segue o padrão PROMOCOES-ATIVAS DD-MM-AAAA.xlsx", () => {
    expect(buildOutputFilename(new Date(2026, 9, 9))).toBe("PROMOCOES-ATIVAS 09-10-2026.xlsx");
    expect(buildOutputFilename(new Date(2027, 0, 2, 23, 59))).toBe("PROMOCOES-ATIVAS 02-01-2027.xlsx");
  });
  it("sem argumento usa a data local de hoje", () => {
    const h = new Date();
    const dd = String(h.getDate()).padStart(2, "0");
    const mm = String(h.getMonth() + 1).padStart(2, "0");
    expect(buildOutputFilename()).toBe(`PROMOCOES-ATIVAS ${dd}-${mm}-${h.getFullYear()}.xlsx`);
  });
});

describe("generateWorkbook", () => {
  it("gera com o nome padrão e a data informada", () => {
    const g = generateWorkbook(texto(3), modelo, new Date(2026, 9, 9));
    expect(g.filename).toBe("PROMOCOES-ATIVAS 09-10-2026.xlsx");
    expect(g.recordCount).toBe(3);
  });

  it("substitui integralmente e ajusta a tabela", () => {
    const out = readOutput(generateWorkbook(texto(5), modelo).data);
    expect(out.rowsWithData()).toEqual([2, 3, 4, 5, 6]);
    expect(out.row(2)[0]).toBe(500);
    expect(out.tableXml).toContain('ref="A1:F6"');
  });

  it("ajusta linhas para mais registros que o modelo", () => {
    const out = readOutput(generateWorkbook(texto(130), modelo).data);
    expect(out.row(131)[0]).toBe(629);
    expect(out.row(132)[0]).toBeNull();
    expect(out.tableXml).toContain('ref="A1:F131"');
  });

  it("não altera o modelo recebido", () => {
    const copia = new Uint8Array(modelo);
    generateWorkbook(texto(10), modelo);
    expect(() => generateWorkbook("lixo", modelo)).toThrow(InvalidInput);
    expect([...modelo]).toEqual([...copia]);
  });

  it.each(["", "   ", "\n\n\t\n"])("texto vazio (%j) pede para colar", (vazio) => {
    expect(() => generateWorkbook(vazio, modelo)).toThrow(/Cole o resultado/);
  });

  it("validação automática informa linha e motivo", () => {
    try {
      generateWorkbook(texto(2) + "abc\tNome\t\t10/03/2025\t2026-12-31\t1", modelo);
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(InvalidInput);
      const lines = (e as InvalidInput).lines;
      expect(lines.some((l) => l.startsWith("Linha 3 [COD]"))).toBe(true);
      expect(lines.some((l) => l.includes("dia/mês/ano"))).toBe(true);
    }
  });

  it("cabeçalho colado impede a geração", () => {
    const cab = "COD\tNOME\tDETALHES\tDATA INICIAL\tDATA FINAL\tPRIORIDADE\n";
    expect(() => generateWorkbook(cab + texto(2), modelo)).toThrow(/cabeçalhos/);
  });

  it("limita a quantidade de erros mostrados a 5", () => {
    const ruim = Array.from({ length: 20 }, () => "x\ty").join("\n");
    try {
      generateWorkbook(ruim, modelo);
      expect.unreachable();
    } catch (e) {
      expect((e as InvalidInput).lines).toHaveLength(5);
      expect((e as InvalidInput).total).toBe(20);
    }
  });

  it("avisos não impedem a geração", () => {
    const g = generateWorkbook("1\tA\tx\t2026-05-01\t2026-04-01\t1\n1\tB\t\t2026-01-01\t2026-02-01\t1", modelo);
    expect(g.recordCount).toBe(2);
    expect(g.warnings.length).toBeGreaterThan(0);
  });

  it("modelo inválido gera mensagem clara", () => {
    expect(() => generateWorkbook(texto(2), new Uint8Array([1, 2, 3, 4, 5]))).toThrow(TemplateError);
    expect(() => generateWorkbook(texto(2), new Uint8Array([1, 2, 3, 4, 5]))).toThrow(/modelo interno/);
  });

  it("todos os erros esperados são ExcelServiceError (mensagem amigável)", () => {
    for (const entrada of ["", "a\tb", "1\tA\tx\t2026-02-30\t2026-12-31\t1"]) {
      expect(() => generateWorkbook(entrada, modelo)).toThrow(ExcelServiceError);
    }
  });

  it("fluxo completo texto -> planilha com acentos e campo vazio", () => {
    const t = "187\tPetit Deli Ciclo 3 - 2025\tGranéis – até R$ 2,87\t2025-03-10\t2026-12-31\t4\n233\tKIT VITALÍCIO\t\t2025-05-05\t2026-12-31\t1\n";
    const out = readOutput(generateWorkbook(t, modelo).data);
    expect(out.row(2)).toEqual([187, "Petit Deli Ciclo 3 - 2025", "Granéis – até R$ 2,87", 45726, 46387, 4]);
    expect(out.row(3)[2]).toBeNull();
    expect(out.row(4)).toEqual([null, null, null, null, null, null]);
  });
});

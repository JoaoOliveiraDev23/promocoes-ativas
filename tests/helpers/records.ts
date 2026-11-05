import type { PromoRecord } from "../../src/core/models";

export function makeRecords(n: number, startCod = 1000): PromoRecord[] {
  return Array.from({ length: n }, (_v, i) => ({
    cod: startCod + i,
    nome: `Promoção ${i} – Ciclo ${i % 13}`,
    detalhes: `Compre ${i + 2} unidades por R$ ${i + 1},99`,
    dataInicial: { year: 2026, month: 1, day: 1 + (i % 28) },
    dataFinal: { year: 2026, month: 12, day: 31 },
    prioridade: 1 + (i % 4),
  }));
}

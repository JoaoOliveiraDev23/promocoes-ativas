import { OUTPUT_PREFIX } from "./constants";

/** `PROMOCOES-ATIVAS DD-MM-AAAA.xlsx` com a data LOCAL de geração. */
export function buildOutputFilename(when: Date = new Date()): string {
  const dd = String(when.getDate()).padStart(2, "0");
  const mm = String(when.getMonth() + 1).padStart(2, "0");
  return `${OUTPUT_PREFIX} ${dd}-${mm}-${when.getFullYear()}.xlsx`;
}

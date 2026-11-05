import { analyzeTemplate, openPackage, type Files } from "../../src/core/xlsx/template";
import { text } from "../../src/core/xlsx/template";
import { parseSharedStrings, parseSheetData, readCellValue, type CellValue } from "../../src/core/xlsx/sheetXml";

export interface Output {
  files: Files;
  sheetXml: string;
  tableXml: string | undefined;
  stylesXml: string;
  sst: string[];
  /** Valores das colunas A..F de uma linha. */
  row(r: number): CellValue[];
  /** Atributo s (estilo) de cada célula A..F da linha. */
  styles(r: number): Array<string | undefined>;
  /** Linhas (>=2) que têm algum VALOR em A:F. */
  rowsWithData(): number[];
  /** Linhas (>=2) que têm qualquer célula em A:F (inclusive vazias formatadas). */
  rowsWithAnyCell(): number[];
}

export function readOutput(bytes: Uint8Array, sheetPath = "xl/worksheets/sheet1.xml"): Output {
  const files = openPackage(bytes);
  const sheetXml = text(files, sheetPath)!;
  const data = parseSheetData(sheetXml)!;
  const sst = parseSharedStrings(text(files, "xl/sharedStrings.xml"));
  const cells = (r: number) => data.rows.find((x) => x.r === r)?.cells ?? [];
  return {
    files,
    sheetXml,
    tableXml: text(files, "xl/tables/table1.xml"),
    stylesXml: text(files, "xl/styles.xml")!,
    sst,
    row: (r) => [1, 2, 3, 4, 5, 6].map((c) => {
      const cell = cells(r).find((x) => x.col === c);
      return cell ? readCellValue(cell, sst) : null;
    }),
    styles: (r) => [1, 2, 3, 4, 5, 6].map((c) => {
      const raw = cells(r).find((x) => x.col === c)?.attrs ?? "";
      return /\ss="(\d+)"/.exec(raw)?.[1];
    }),
    rowsWithData: () =>
      data.rows.filter((x) => x.r >= 2 && x.cells.some((c) => c.col <= 6 && readCellValue(c, sst) !== null)).map((x) => x.r),
    rowsWithAnyCell: () => data.rows.filter((x) => x.r >= 2 && x.cells.some((c) => c.col <= 6)).map((x) => x.r),
  };
}

export { analyzeTemplate };

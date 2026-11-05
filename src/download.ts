export const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export function createDownloadUrl(data: Uint8Array): string {
  return URL.createObjectURL(new Blob([data as BlobPart], { type: XLSX_MIME }));
}

/** Inicia o download no navegador (chamado dentro do clique do usuário). */
export function startDownload(url: string, filename: string): void {
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
}

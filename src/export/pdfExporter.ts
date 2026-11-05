
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import type { PromoRecord } from "../core/models";

/**
 * Formata uma data de calendário sem utilizar o objeto Date.
 * Mantém o padrão AAAA-MM-DD utilizado pelo sistema.
 */
function formatDate(date: PromoRecord["dataInicial"]): string {
  const year = String(date.year);
  const month = String(date.month).padStart(2, "0");
  const day = String(date.day).padStart(2, "0");

  return `${day}-${month}-${year}`;
}

/**
 * Gera e baixa um PDF contendo os registros de promoções.
 */
export function exportPromotionsPdf(
  records: PromoRecord[],
  filename: string
): void {
  if (records.length === 0) {
    throw new Error("Não existem registros para exportar.");
  }

  const doc = new jsPDF({
    orientation: "landscape",
    unit: "mm",
    format: "a4",
  });

  const margin = {
    top: 22,
    right: 10,
    bottom: 15,
    left: 10,
  };

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();

  const body = records.map((record) => [
    String(record.cod),
    record.nome,
    record.detalhes || "-",
    formatDate(record.dataInicial),
    formatDate(record.dataFinal),
    String(record.prioridade),
  ]);

  autoTable(doc, {
    head: [[
      "COD",
      "NOME",
      "DETALHES",
      "DATA INICIAL",
      "DATA FINAL",
      "PRIORIDADE",
    ]],

    body,

    startY: margin.top,

    margin,

    tableWidth: "auto",

    theme: "grid",

    styles: {
      font: "helvetica",
      fontSize: 8,
      cellPadding: 2,
      overflow: "linebreak",
      valign: "middle",
      lineColor: [210, 210, 210],
      lineWidth: 0.15,
    },

    headStyles: {
      fillColor: [35, 35, 35],
      textColor: [255, 255, 255],
      fontStyle: "bold",
      halign: "center",
      valign: "middle",
    },

    alternateRowStyles: {
      fillColor: [245, 245, 245],
    },

    columnStyles: {
      0: { cellWidth: 25, halign: "center" },
      1: { cellWidth: 40 },
      2: { cellWidth: 120 },
      3: { cellWidth: 30, halign: "center" },
      4: { cellWidth: 30, halign: "center" },
      5: { cellWidth: 25, halign: "center" },
    },

    didDrawPage: (data) => {
      // Cabeçalho
      doc.setFont("helvetica", "bold");
      doc.setFontSize(14);
      doc.setTextColor(35, 35, 35);
      doc.text("Promoções Ativas", margin.left, 10);

      doc.setFont("helvetica", "normal");
      doc.setFontSize(8);
      doc.setTextColor(110, 110, 110);
      doc.text(
        `Consulta realizada em: ${new Date().toLocaleDateString("pt-BR")}`,
        margin.left,
        16
      );

      // Rodapé
      doc.setFontSize(8);
      doc.text(
        `Página ${data.pageNumber}`,
        pageWidth - margin.right,
        pageHeight - 5,
        { align: "right" }
      );
    },
  });

  // Substitui a extensão XLSX pela extensão PDF.
  const pdfFilename = filename.replace(/\.xlsx$/i, ".pdf");

  doc.save(pdfFilename);
}
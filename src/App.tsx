import { useEffect, useRef, useState } from "react";
import { Message } from "./components/Message";
import { generateWorkbook, InvalidInput } from "./core/generator";
import { ExcelServiceError } from "./core/xlsx";
import { createDownloadUrl, startDownload } from "./download";
import { loadTemplate } from "./loadTemplate";

type TemplateState =
  | { status: "loading" }
  | { status: "ready"; bytes: Uint8Array }
  | { status: "error"; message: string };

type Feedback =
  | { kind: "success"; filename: string; count: number; url: string; warnings: string[] }
  | { kind: "error"; heading?: string; items: string[] }
  | null;

export default function App() {
  const [text, setText] = useState("");
  const [template, setTemplate] = useState<TemplateState>({ status: "loading" });
  const [feedback, setFeedback] = useState<Feedback>(null);
  const currentUrl = useRef<string | null>(null);

  // Carrega o modelo interno uma vez.
  useEffect(() => {
    let cancelled = false;
    loadTemplate()
      .then((bytes) => !cancelled && setTemplate({ status: "ready", bytes }))
      .catch((err: unknown) => {
        if (!cancelled) setTemplate({ status: "error", message: err instanceof Error ? err.message : "Erro ao carregar o modelo." });
      });
    return () => {
      cancelled = true;
    };
  }, []);

  // Libera o arquivo temporário do navegador ao sair.
  useEffect(() => () => releaseUrl(), []);

  function releaseUrl() {
    if (currentUrl.current) URL.revokeObjectURL(currentUrl.current);
    currentUrl.current = null;
  }

  function handleChange(value: string) {
    setText(value);
    if (feedback) {
      releaseUrl();
      setFeedback(null); // editar o texto limpa a mensagem anterior
    }
  }

  function handleGenerate() {
    if (template.status !== "ready") return;
    releaseUrl();
    try {
      const result = generateWorkbook(text, template.bytes);
      const url = createDownloadUrl(result.data);
      currentUrl.current = url;
      startDownload(url, result.filename);
      setFeedback({ kind: "success", filename: result.filename, count: result.recordCount, url, warnings: result.warnings });
    } catch (exc) {
      if (exc instanceof InvalidInput) {
        const items = [...exc.lines];
        if (exc.total > exc.lines.length) items.push(`… e mais ${exc.total - exc.lines.length} erro(s).`);
        setFeedback({ kind: "error", heading: "Não foi possível gerar a planilha:", items });
      } else if (exc instanceof ExcelServiceError) {
        setFeedback({ kind: "error", items: [exc.message] });
      } else {
        console.error("Falha inesperada na geração:", exc instanceof Error ? exc.name : "erro");
        setFeedback({ kind: "error", items: ["Ocorreu um erro inesperado ao gerar a planilha."] });
      }
    }
  }

  return (
    <main className="page">
      <h1 className="title">Atualizador de Promoções</h1>

      {template.status === "error" ? (
        <Message kind="error">{template.message}</Message>
      ) : (
        <>
          <textarea
            className="input"
            aria-label="Resultado da consulta SQL Server"
            placeholder="Cole aqui o resultado da consulta SQL Server"
            value={text}
            onChange={(e) => handleChange(e.target.value)}
            spellCheck={false}
            wrap="off"
          />

          <button className="button" type="button" onClick={handleGenerate} disabled={template.status !== "ready"}>
            Gerar e baixar planilha
          </button>

          {feedback?.kind === "error" && (
            <Message kind="error">
              {feedback.heading && <p>{feedback.heading}</p>}
              {feedback.items.length === 1 && !feedback.heading ? (
                <p>{feedback.items[0]}</p>
              ) : (
                <ul>
                  {feedback.items.map((item, i) => (
                    <li key={i}>{item}</li>
                  ))}
                </ul>
              )}
            </Message>
          )}

          {feedback?.kind === "success" && (
            <>
              <Message kind="success">
                Planilha gerada com {feedback.count} registro(s): {feedback.filename}
              </Message>
              {feedback.warnings.length > 0 && (
                <Message kind="warning">
                  <p>Avisos (não impediram a geração):</p>
                  <ul>
                    {feedback.warnings.map((w, i) => (
                      <li key={i}>{w}</li>
                    ))}
                  </ul>
                </Message>
              )}
              <a className="fallback" href={feedback.url} download={feedback.filename}>
                Se o download não iniciar, clique aqui.
              </a>
            </>
          )}
        </>
      )}
    </main>
  );
}

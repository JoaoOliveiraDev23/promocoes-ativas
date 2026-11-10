
import { useEffect, useRef, useState } from "react";
import { Message } from "./components/Message";
import {
  generateWorkbook,
  validatePromotions,
  InvalidInput,
} from "./core/generator";
import { buildPdfFilename } from "./core/filename";
import { ExcelServiceError } from "./core/xlsx";
import { createDownloadUrl, startDownload } from "./download";
import { exportPromotionsPdf } from "./export/pdfExporter";
import { loadTemplate } from "./loadTemplate";

type TemplateState =
  | { status: "loading" }
  | { status: "ready"; bytes: Uint8Array }
  | { status: "error"; message: string };

type Feedback =
  | {
      kind: "success";
      id: number;
      format: "excel" | "pdf";
      filename: string;
      count: number;
      url?: string;
      warnings: string[];
    }
  | {
      kind: "error";
      id: number;
      format: "planilha" | "PDF";
      heading?: string;
      items: string[];
    }
  | null;

export default function App() {
  const [text, setText] = useState("");
  const [template, setTemplate] = useState<TemplateState>({
    status: "loading",
  });
  const [feedback, setFeedback] = useState<Feedback>(null);

  const currentUrl = useRef<string | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Identificador individual de cada tentativa de geração.
  const feedbackId = useRef(0);

  // Tema --------------------------------------------------------------------

  // Recupera a preferência de tema salva no navegador.
  // Caso não exista uma preferência anterior, utiliza o tema escuro.
  const [theme, setTheme] = useState<"light" | "dark">(() => {
    const savedTheme = localStorage.getItem("theme");

    return savedTheme === "light" ? "light" : "dark";
  });

  // Aplica o tema selecionado e salva a preferência no navegador.
  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem("theme", theme);
  }, [theme]);

  // Alterna entre o tema claro e o tema escuro.
  function handleThemeToggle() {
    setTheme((currentTheme) =>
      currentTheme === "light" ? "dark" : "light"
    );
  }

  // Modelo interno ----------------------------------------------------------

  // Carrega o modelo interno uma vez.
  useEffect(() => {
    let cancelled = false;

    loadTemplate()
      .then((bytes) => {
        if (!cancelled) {
          setTemplate({ status: "ready", bytes });
        }
      })
      .catch((err: unknown) => {
        if (!cancelled) {
          setTemplate({
            status: "error",
            message:
              err instanceof Error
                ? err.message
                : "Erro ao carregar o modelo.",
          });
        }
      });

    return () => {
      cancelled = true;
    };
  }, []);

  // Download ----------------------------------------------------------------

  // Libera o arquivo temporário do navegador ao sair.
  useEffect(() => {
    return () => {
      if (currentUrl.current) {
        URL.revokeObjectURL(currentUrl.current);
      }
    };
  }, []);

  function releaseUrl() {
    if (currentUrl.current) {
      URL.revokeObjectURL(currentUrl.current);
    }

    currentUrl.current = null;
  }

  // Controle das mensagens --------------------------------------------------

  // Inicia uma nova tentativa de geração.
  // Encerra qualquer feedback anterior e libera URLs temporárias.
  function beginGeneration() {
    feedbackId.current += 1;

    releaseUrl();
    setFeedback(null);

    return feedbackId.current;
  }

  // Mensagens de erro -------------------------------------------------------

  // Remove mensagens de erro após 5 segundos.
  useEffect(() => {
    if (feedback?.kind !== "error") return;

    const currentId = feedback.id;

    const timeout = window.setTimeout(() => {
      setFeedback((current) =>
        current?.kind === "error" && current.id === currentId
          ? null
          : current
      );
    }, 5000);

    return () => {
      window.clearTimeout(timeout);
    };
  }, [feedback]);

  // Entrada de dados --------------------------------------------------------

  function handleChange(value: string) {
    setText(value);

    if (feedback) {
      // Invalida a mensagem anterior.
      feedbackId.current += 1;

      releaseUrl();
      setFeedback(null);
    }
  }

  // Posiciona o cursor no início da última linha após colar.
  function handlePaste() {
    requestAnimationFrame(() => {
      const textarea = textareaRef.current;

      if (!textarea) return;

      const value = textarea.value;
      const lastLineStart = value.lastIndexOf("\n") + 1;

      textarea.focus();
      textarea.setSelectionRange(lastLineStart, lastLineStart);

      // Leva a rolagem horizontal e vertical até a posição do cursor.
      const cursorPosition = textarea.selectionStart;

      // Calcula a posição vertical do cursor.
      const textBeforeCursor = value.slice(0, cursorPosition);
      const linesBeforeCursor = textBeforeCursor.split("\n").length;

      const lineHeight = parseFloat(
        getComputedStyle(textarea).lineHeight
      );

      textarea.scrollTop = Math.max(
        0,
        (linesBeforeCursor - 1) * lineHeight
      );

      // Posiciona a rolagem horizontal no início da linha.
      textarea.scrollLeft = 0;
    });
  }



  // Tratamento compartilhado de erros --------------------------------------

  function handleGenerationError(
    exc: unknown,
    format: "planilha" | "PDF",
    attemptId: number
  ) {
    if (exc instanceof InvalidInput) {
      const items = [...exc.lines];

      if (exc.total > exc.lines.length) {
        items.push(`… e mais ${exc.total - exc.lines.length} erro(s).`);
      }

      setFeedback({
        kind: "error",
        id: attemptId,
        format,
        heading: `Não foi possível gerar ${format}:`,
        items,
      });
    } else if (exc instanceof ExcelServiceError) {
      setFeedback({
        kind: "error",
        id: attemptId,
        format,
        items: [exc.message],
      });
    } else {
      console.error(
        `Falha inesperada na geração do ${format}:`,
        exc instanceof Error ? exc.name : "erro"
      );

      setFeedback({
        kind: "error",
        id: attemptId,
        format,
        items: [
          `Ocorreu um erro inesperado ao gerar o ${format}.`,
        ],
      });
    }
  }

  // Geração da planilha -----------------------------------------------------

  function handleGenerate() {
    const attemptId = beginGeneration();

    if (template.status !== "ready") return;

    try {
      const result = generateWorkbook(text, template.bytes);
      const url = createDownloadUrl(result.data);

      currentUrl.current = url;

      startDownload(url, result.filename);

      setFeedback({
        kind: "success",
        id: attemptId,
        format: "excel",
        filename: result.filename,
        count: result.recordCount,
        url,
        warnings: result.warnings,
      });
    } catch (exc) {
      handleGenerationError(exc, "planilha", attemptId);
    }
  }

  // Geração do PDF ----------------------------------------------------------

  function handleGeneratePdf() {
    const attemptId = beginGeneration();

    try {
      const result = validatePromotions(text);
      const filename = buildPdfFilename();

      exportPromotionsPdf(result.records, filename);

      setFeedback({
        kind: "success",
        id: attemptId,
        format: "pdf",
        filename,
        count: result.records.length,
        warnings: result.warnings,
      });
    } catch (exc) {
      handleGenerationError(exc, "PDF", attemptId);
    }
  }

  // Interface ---------------------------------------------------------------

  return (
    <main className="page">

      {/* Cabeçalho e alternância de tema */}
      <header className="header">
        <h1 className="title">Atualizador de Promoções</h1>

        <button
          className="theme-toggle"
          type="button"
          onClick={handleThemeToggle}
          aria-label={
            theme === "light"
              ? "Ativar tema escuro"
              : "Ativar tema claro"
          }
          title={
            theme === "light"
              ? "Ativar tema escuro"
              : "Ativar tema claro"
          }
        >
          <span className="theme-icon-viewport" aria-hidden="true">
            <span
              className={`theme-icon theme-icon-moon ${
                theme === "dark" ? "is-visible" : "is-hidden"
              }`}
            >
              🌙
            </span>

            <span
              className={`theme-icon theme-icon-sun ${
                theme === "light" ? "is-visible" : "is-hidden"
              }`}
            >
              ☀️
            </span>
          </span>
        </button>
      </header>

      {/* Aviso de carregamento do modelo */}
      {template.status === "error" && (
        <Message kind="error">
          {template.message}
          <p>O download em PDF continua disponível.</p>
        </Message>
      )}

      {/* Campo para inserir os dados da consulta SQL Server */}
      <textarea
        ref={textareaRef}
        className="input"
        id="input"
        aria-label="Resultado da consulta SQL Server"
        placeholder="Cole aqui o resultado da consulta SQL Server"
        value={text}
        onChange={(e) => {
          handleChange(e.target.value);

          const textarea = e.target;

          // No celular, mantém a altura definida pelo CSS.
          if (window.matchMedia("(max-width: 600px)").matches) {
            textarea.style.height = "";
            textarea.style.overflowY = "auto";
            return;
          }

          // No desktop, permite o crescimento automático.
          textarea.style.height = "auto";

          const maxHeight = parseFloat(
            getComputedStyle(textarea).maxHeight
          );

          const contentHeight = textarea.scrollHeight;

          // Limita o crescimento à altura máxima definida no CSS.
          textarea.style.height = `${Math.min(
            contentHeight,
            maxHeight
          )}px`;

          // Exibe a rolagem vertical somente quando necessário.
          textarea.style.overflowY =
            contentHeight > maxHeight ? "auto" : "hidden";
        }}
        onPaste={handlePaste}
        spellCheck={false}
        wrap="off"
        rows={1}
      />

      {/* Botões de geração e download */}
      <div className="actions">
        <button
          className="button"
          type="button"
          onClick={handleGenerate}
          disabled={template.status !== "ready"}
        >
          Exportar em Excel
        </button>

        <button
          className="button"
          type="button"
          onClick={handleGeneratePdf}
        >
          Exportar em PDF
        </button>
      </div>

      {/* Exibição de mensagens de erro */}
      {feedback?.kind === "error" && (
        <div
          className="error-feedback"
          key={feedback.id}
        >
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
        </div>
      )}

      {/* Exibição de mensagens de sucesso e avisos */}
      {feedback?.kind === "success" && (
        <>
          <Message kind="success">
            Arquivo{" "}
            {feedback.format === "excel" ? "Excel" : "PDF"} gerado
            com {feedback.count} registro(s): {feedback.filename}
          </Message>

          {/* Avisos que não impediram a geração */}
          {feedback.warnings.length > 0 && (
            <Message kind="warning">
              <p>Avisos (não impediram a geração):</p>

              <ul>
                {feedback.warnings.map((warning, i) => (
                  <li key={i}>{warning}</li>
                ))}
              </ul>
            </Message>
          )}

          {/* Link alternativo para download da planilha */}
          {feedback.format === "excel" && feedback.url && (
            <a
              className="fallback"
              href={feedback.url}
              download={feedback.filename}
            >
              Se o download não iniciar, clique aqui.
            </a>
          )}
        </>
      )}
    </main>
  );
}
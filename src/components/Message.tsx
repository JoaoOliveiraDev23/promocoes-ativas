import type { ReactNode } from "react";

interface MessageProps {
  kind: "success" | "error" | "warning";
  children: ReactNode;
}

/** Caixa de mensagem (sucesso, erro ou aviso). Estilos em styles.css (.message). */
export function Message({ kind, children }: MessageProps) {
  return (
    <div className={`message message--${kind}`} role={kind === "error" ? "alert" : "status"}>
      {children}
    </div>
  );
}

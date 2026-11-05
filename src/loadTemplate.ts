import { TEMPLATE_DISPLAY_PATH, TEMPLATE_FILE_NAME } from "./core/constants";
import { TemplateError } from "./core/xlsx";

/**
 * Carrega o modelo interno (public/PROMOCOES-ATIVAS-modelo.xlsx), servido como arquivo estático.
 * O arquivo só é lido; nunca é enviado nem alterado.
 */
export async function loadTemplate(): Promise<Uint8Array> {
  const missing = new TemplateError(
    `O arquivo de modelo interno não foi encontrado: ${TEMPLATE_DISPLAY_PATH}. ` +
      "Restaure-o na pasta do projeto e publique novamente.",
  );

  let response: Response;
  try {
    response = await fetch(`${import.meta.env.BASE_URL}${TEMPLATE_FILE_NAME}`, { cache: "no-cache" });
  } catch {
    throw new TemplateError(
      `Não foi possível carregar o modelo interno (${TEMPLATE_DISPLAY_PATH}). Verifique a conexão e recarregue a página.`,
    );
  }
  if (!response.ok) throw missing;

  const bytes = new Uint8Array(await response.arrayBuffer());
  // Alguns servidores respondem 200 com a página inicial quando o arquivo não existe.
  if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw missing;
  return bytes;
}

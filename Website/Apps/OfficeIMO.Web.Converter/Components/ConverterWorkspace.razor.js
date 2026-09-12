export function createObjectUrl(bytes, contentType) {
  const blob = new Blob([bytes], { type: contentType || "application/octet-stream" });
  return URL.createObjectURL(blob);
}

const toolName = "convert_selected_document";
let activeConverter = null;
let registered = false;
let registrationController = null;

function toolDocument() {
  // Keep the tool discoverable on the canonical workspace while execution stays in this app.
  try {
    if (window.parent !== window && window.frameElement?.matches('iframe[data-workspace-src]') &&
        window.parent.location.origin === window.location.origin) return window.parent.document;
  } catch { /* Other-origin embeds retain their own document context. */ }
  return document;
}

export async function registerWebMcpTool(converter) {
  activeConverter = converter;
  const contextDocument = toolDocument();
  if (registered || !contextDocument.modelContext || typeof contextDocument.modelContext.registerTool !== "function") {
    document.body.setAttribute("data-webmcp-status", registered ? "registered" : "unsupported");
    return false;
  }

  const controller = new AbortController();
  registrationController = controller;
  try {
    await contextDocument.modelContext.registerTool({
      name: toolName,
      description: "Convert the document already selected in the visible OfficeIMO workspace using the current browser-local route and settings.",
      inputSchema: {
        type: "object",
        properties: {},
        additionalProperties: false
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
        untrustedContentHint: true
      },
      execute: async (_input, context) => {
        if (!activeConverter) {
          throw new Error("The OfficeIMO converter is no longer available on this page.");
        }
        if (context?.signal?.aborted) {
          return {
            success: false,
            message: "Conversion was cancelled before it started."
          };
        }
        return activeConverter.invokeMethodAsync("ConvertSelectedDocumentForWebMcpAsync");
      }
    }, { signal: controller.signal });
    if (controller.signal.aborted) return false;
    registered = true;
    document.body.setAttribute("data-webmcp-status", "registered");
    return true;
  } catch {
    controller.abort();
    if (registrationController === controller) registrationController = null;
    document.body.setAttribute("data-webmcp-status", "failed");
    return false;
  }
}

export async function unregisterWebMcpTool() {
  activeConverter = null;
  registrationController?.abort();
  registrationController = null;
  registered = false;
  document.body.setAttribute("data-webmcp-status", "disposed");
}

window.addEventListener('pagehide', event => {
  registrationController?.abort();
  registrationController = null;
  registered = false;
  if (!event.persisted) activeConverter = null;
});
window.addEventListener('pageshow', event => {
  if (event.persisted && activeConverter) return registerWebMcpTool(activeConverter);
});

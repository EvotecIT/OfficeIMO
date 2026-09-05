export function createObjectUrl(bytes, contentType) {
  const blob = new Blob([bytes], { type: contentType || "application/octet-stream" });
  return URL.createObjectURL(blob);
}

export function revokeObjectUrl(url) {
  if (url) {
    URL.revokeObjectURL(url);
  }
}

const toolName = "convert_selected_document";
let activeConverter = null;
let registered = false;
let registrationController = null;

export async function registerWebMcpTool(converter) {
  activeConverter = converter;
  if (registered || !document.modelContext || typeof document.modelContext.registerTool !== "function") {
    document.body.setAttribute("data-webmcp-status", registered ? "registered" : "unsupported");
    return false;
  }

  try {
    registrationController = new AbortController();
    await document.modelContext.registerTool({
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
    }, { signal: registrationController.signal });
    registered = true;
    document.body.setAttribute("data-webmcp-status", "registered");
    return true;
  } catch {
    registrationController?.abort();
    registrationController = null;
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

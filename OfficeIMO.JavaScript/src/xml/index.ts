import { OfficeIMOError } from "../core/errors.js";
import { ChunkedTextSink } from "../core/sinks.js";
import type { ByteSink } from "../core/sinks.js";

export type InvalidCharacterPolicy = "strip" | "reject";
export const xmlDeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** XML 1.0 characters; Unicode mode retains valid surrogate pairs. */
export function cleanXml(value: unknown, policy: InvalidCharacterPolicy = "strip"): string {
  if (policy !== "strip" && policy !== "reject") throw new TypeError("Invalid XML character policy.");
  const text = String(value), invalid = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/gu;
  if (policy === "reject" && invalid.test(text)) throw new OfficeIMOError("INVALID_XML", "Text contains XML 1.0-invalid characters.");
  return text.replace(invalid, "");
}

/** Escape data, including whitespace that XML parsers otherwise normalize. */
export function escapeXml(value: unknown, policy: InvalidCharacterPolicy = "strip"): string {
  const escapes: Record<string, string> = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;",
    "\r": "&#13;", "\n": "&#10;", "\t": "&#9;" };
  return cleanXml(value, policy).replace(/[&<>"'\r\n\t]/g, c => escapes[c]!);
}

/** Names are schema-owned ASCII QNames; data is accepted only as text or attribute values. */
export function validateXmlName(name: string): string {
  if (typeof name !== "string" || !/^[A-Za-z_][A-Za-z0-9_.-]*(?::[A-Za-z_][A-Za-z0-9_.-]*)?$/.test(name))
    throw new OfficeIMOError("INVALID_XML", "Invalid XML name: " + name);
  return name;
}

export class XmlWriter {
  private readonly sink: ChunkedTextSink;
  private readonly stack: string[] = [];
  private roots = 0;
  private closed = false;
  constructor(sink: ByteSink, private readonly policy: InvalidCharacterPolicy = "strip", signal?: AbortSignal) {
    cleanXml("", policy);
    this.sink = new ChunkedTextSink(sink, signal);
    this.sink.append(xmlDeclaration);
  }
  async startElement(name: string, attributes: Readonly<Record<string, unknown>> = {}): Promise<void> {
    this.open(); validateXmlName(name);
    const text = "<" + name + Object.entries(attributes).map(([key, value]) => " " + validateXmlName(key) + '="' + escapeXml(value, this.policy) + '"').join("") + ">";
    if (!this.stack.length && this.roots++) throw new OfficeIMOError("INVALID_XML", "XML must have one root element.");
    this.stack.push(name); await this.sink.write(text);
  }
  async text(value: unknown): Promise<void> {
    this.open();
    if (!this.stack.length) throw new OfficeIMOError("INVALID_XML", "Text needs an open element.");
    await this.sink.write(escapeXml(value, this.policy));
  }
  async endElement(): Promise<void> {
    this.open();
    const name = this.stack.pop();
    if (!name) throw new OfficeIMOError("INVALID_XML", "No open XML element.");
    await this.sink.write("</" + name + ">");
  }
  async close(): Promise<void> {
    this.open();
    if (this.stack.length || this.roots !== 1) throw new OfficeIMOError("INVALID_XML", "XML document is incomplete.");
    await this.sink.close(); this.closed = true;
  }
  private open(): void { if (this.closed) throw new OfficeIMOError("INVALID_STATE", "XML writer is closed."); }
}

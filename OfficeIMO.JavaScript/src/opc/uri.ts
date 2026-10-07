import { OfficeIMOError } from "../core/errors.js";

/** Canonical absolute OPC part URI. ASCII URI spelling; Unicode must be UTF-8 percent encoded. */
export function partUri(value: string): string {
  const bad = () => new OfficeIMOError("INVALID_PART_URI", "Invalid OPC part URI: " + value);
  if (typeof value !== "string" || !value.startsWith("/") || /[^A-Za-z0-9\-._~!$&'()*+,;=:@%/]/.test(value)) throw bad();
  if (value.split("/").slice(1).some(s => !s || s.endsWith(".") || /^\.+$/.test(s))) throw bad();
  if (/%(?![\da-f]{2})/i.test(value)) throw bad();
  for (const match of value.matchAll(/%([\da-f]{2})/gi)) {
    const char = String.fromCharCode(parseInt(match[1]!, 16));
    if (/[A-Za-z0-9_.~\-/\\]/.test(char)) throw bad();
  }
  try { decodeURIComponent(value); } catch { throw bad(); }
  return value.replace(/%[\da-f]{2}/gi, s => s.toUpperCase());
}

export function relationshipPartUri(source: string): string {
  if (source === "/") return "/_rels/.rels";
  const uri = partUri(source), slash = uri.lastIndexOf("/");
  return uri.slice(0, slash + 1) + "_rels/" + uri.slice(slash + 1) + ".rels";
}

/** Relative target, derived from validated part URIs rather than caller-provided traversal. */
export function relativePartTarget(source: string, target: string): string {
  const to = partUri(target).slice(1).split("/");
  const from = source === "/" ? [] : partUri(source).slice(1).split("/").slice(0, -1);
  while (from.length && from[0] === to[0]) { from.shift(); to.shift(); }
  const relative = "../".repeat(from.length) + to.join("/");
  // RFC 3986 section 4.2 forbids a colon in a relative reference's first segment.
  return relative.split("/", 1)[0]!.includes(":") ? "./" + relative : relative;
}

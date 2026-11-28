import { NotSupportedError } from "../../core/errors.js";

/** One parsed SVG element. Names keep their prefix; `local` is the name after it. */
export interface SvgNode {
  readonly name: string;
  readonly local: string;
  readonly attributes: ReadonlyMap<string, string>;
  readonly children: readonly (SvgNode | string)[];
}

const attributePattern = /([A-Za-z_][\w.:-]*)\s*=\s*("([^"]*)"|'([^']*)')/y;
const entities: Readonly<Record<string, string>> = { lt: "<", gt: ">", amp: "&", quot: "\"", apos: "'", nbsp: " " };

function decode(text: string): string {
  return text.indexOf("&") < 0 ? text : text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[A-Za-z]+);/g, (match, entity: string) => {
    if (entity[0] === "#") {
      const code = entity[1] === "x" || entity[1] === "X" ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10);
      return code > 0 && code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff) ? String.fromCodePoint(code) : "�";
    }
    return entities[entity] ?? match;
  });
}

/** A small, non-validating parser for serialized SVG: elements, attributes, text, CDATA and character references.
 * Comments, processing instructions and doctype declarations are skipped; external entities are never resolved. */
export function parseSvgMarkup(markup: string, maxDepth = 256): SvgNode {
  if (typeof markup !== "string") throw new TypeError("Chart SVG must be serialized markup.");
  type Open = { name: string; local: string; attributes: Map<string, string>; children: (SvgNode | string)[] };
  const stack: Open[] = [];
  let root: SvgNode | undefined, index = 0;
  const text = (value: string): void => { const top = stack[stack.length - 1]; if (top && value) top.children.push(value); };
  while (index < markup.length) {
    const open = markup.indexOf("<", index);
    if (open < 0) { text(decode(markup.slice(index))); break; }
    if (open > index) text(decode(markup.slice(index, open)));
    if (markup.startsWith("<!--", open)) {
      const end = markup.indexOf("-->", open + 4);
      if (end < 0) throw new TypeError("Chart SVG has an unterminated comment.");
      index = end + 3; continue;
    }
    if (markup.startsWith("<![CDATA[", open)) {
      const end = markup.indexOf("]]>", open + 9);
      if (end < 0) throw new TypeError("Chart SVG has an unterminated CDATA section.");
      text(markup.slice(open + 9, end)); index = end + 3; continue;
    }
    if (markup[open + 1] === "?") {
      const end = markup.indexOf("?>", open + 2);
      if (end < 0) throw new TypeError("Chart SVG has an unterminated processing instruction.");
      index = end + 2; continue;
    }
    if (markup[open + 1] === "!") {
      if (/^<!DOCTYPE[^>]*\[/i.test(markup.slice(open, open + 1024))) throw new NotSupportedError("Chart SVG with an internal DTD subset is not supported.");
      const end = markup.indexOf(">", open + 2);
      if (end < 0) throw new TypeError("Chart SVG has an unterminated declaration.");
      index = end + 1; continue;
    }
    if (markup[open + 1] === "/") {
      const end = markup.indexOf(">", open + 2);
      if (end < 0) throw new TypeError("Chart SVG has an unterminated end tag.");
      const name = markup.slice(open + 2, end).trim(), top = stack.pop();
      if (!top || top.name !== name) throw new TypeError("Chart SVG end tag </" + name + "> does not match its element.");
      const node: SvgNode = top;
      if (stack.length) stack[stack.length - 1]!.children.push(node); else root = node;
      index = end + 1; continue;
    }
    const tag = /^<([A-Za-z_][\w.:-]*)/.exec(markup.slice(open, open + 256));
    if (!tag) throw new TypeError("Chart SVG has an invalid start tag.");
    const name = tag[1]!, attributes = new Map<string, string>();
    let position = open + tag[0].length, selfClosing = false;
    for (;;) {
      while (position < markup.length && /\s/.test(markup[position]!)) position++;
      const character = markup[position];
      if (character === undefined) throw new TypeError("Chart SVG has an unterminated start tag.");
      if (character === ">") { position++; break; }
      if (character === "/" && markup[position + 1] === ">") { position += 2; selfClosing = true; break; }
      attributePattern.lastIndex = position;
      const attribute = attributePattern.exec(markup);
      if (!attribute) throw new TypeError("Chart SVG has an invalid attribute on <" + name + ">.");
      attributes.set(attribute[1]!, decode(attribute[3] ?? attribute[4] ?? ""));
      position = attributePattern.lastIndex;
    }
    if (root) throw new TypeError("Chart SVG must have one root element.");
    const local = name.slice(name.indexOf(":") + 1), element: Open = { name, local, attributes, children: [] };
    if (selfClosing) {
      if (stack.length) stack[stack.length - 1]!.children.push(element); else root = element;
    } else {
      if (stack.length >= maxDepth) throw new NotSupportedError("Chart SVG nesting exceeds " + maxDepth + " levels.");
      stack.push(element);
    }
    index = position;
  }
  if (stack.length) throw new TypeError("Chart SVG has an unclosed <" + stack[stack.length - 1]!.name + "> element.");
  if (!root || root.local !== "svg") throw new TypeError("Chart SVG must have an <svg> root element.");
  return root;
}

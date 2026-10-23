import assert from "node:assert/strict";

/** Select actual structure dictionaries, excluding stream payloads that contain packed members. */
export function pdfStructureElements(pdf) {
  return [...pdf.objects].filter(([, object]) => object.stream === undefined && /^\s*<<\s*\/Type\s+\/StructElem\b/.test(object.body));
}

// Narrow semantic reader over the independent xref reader's object bodies. It keeps
// child references distinct from page-local MCIDs and consumes every K value.
export function pdfStructureChildren(body) {
  let position = 0;
  function whitespace() {
    while (position < body.length) {
      if (/\s/.test(body[position])) position++;
      else if (body[position] === "%") while (position < body.length && !/[\r\n]/.test(body[position])) position++;
      else break;
    }
  }
  function token() {
    whitespace(); const start = position;
    for (const delimiter of ["<<", ">>", "[", "]"]) if (body.startsWith(delimiter, position)) {
      position += delimiter.length; return { text: delimiter, start, end: position };
    }
    if (body[position] === "(") {
      let depth = 1; position++;
      while (position < body.length && depth) {
        const character = body[position++];
        if (character === "\\") { assert.ok(position < body.length, "Incomplete PDF string escape"); position++; }
        else if (character === "(") depth++;
        else if (character === ")") depth--;
      }
      assert.equal(depth, 0, "Unterminated PDF string");
    } else if (body[position] === "<") {
      const end = body.indexOf(">", ++position); assert.ok(end >= position, "Unterminated PDF hex string");
      assert.match(body.slice(position, end), /^[\da-f\s]*$/i); position = end + 1;
    } else {
      if (body[position] === "/") position++;
      while (position < body.length && !/[\s<>\[\](){}\/%]/.test(body[position])) position++;
      assert.ok(position > start && body.slice(start, position) !== "/", "Unexpected PDF token");
    }
    return { text: body.slice(start, position), start, end: position };
  }
  let lookahead;
  const peek = () => lookahead ??= token();
  const next = () => { const value = peek(); lookahead = undefined; return value; };
  function value() {
    const first = next(), { text, start } = first;
    if (text === "<<") {
      const fields = new Map();
      while (peek().text !== ">>") {
        const key = next().text; assert.ok(key.startsWith("/"), "A PDF dictionary key must be a name");
        assert.ok(!fields.has(key), "Duplicate PDF dictionary key"); fields.set(key, value());
      }
      return { kind: "dictionary", fields, start, end: next().end };
    }
    if (text === "[") {
      const values = [];
      while (peek().text !== "]") values.push(value());
      return { kind: "array", values, start, end: next().end };
    }
    if (/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(text)) {
      const number = Number(text), integer = /^[+-]?\d+$/.test(text);
      assert.ok(Number.isFinite(number), "Nonfinite PDF number");
      // Only an integer followed by another integer can begin an indirect reference.
      if (integer && /^[+-]?\d+$/.test(peek().text)) {
        const saved = lookahead, savedPosition = position; next();
        if (peek().text === "R") return { kind: "reference", id: number, generation: Number(saved.text), start, end: next().end };
        lookahead = saved; position = savedPosition;
      }
      return { kind: "number", number, integer, start, end: first.end };
    }
    assert.ok(text.startsWith("/") || text.startsWith("(") || text.startsWith("<") || ["true", "false", "null"].includes(text), "Unsupported PDF value");
    return { kind: "literal", text, start, end: first.end };
  }
  function reference(node, label) {
    assert.equal(node?.kind, "reference", label + " must be an object reference");
    assert.ok(Number.isSafeInteger(node.id) && node.id > 0, label + " must name a positive object");
    assert.equal(node.generation, 0, label + " must use generation zero"); return node.id;
  }
  function mark(node, pageId, form) {
    assert.ok(Number.isSafeInteger(pageId) && pageId > 0, "A marked-content child requires a default or explicit page");
    assert.ok(node?.kind === "number" && node.integer && Number.isSafeInteger(node.number) && node.number >= 0, "MCID must be a nonnegative safe integer");
    return { kind: "mark", pageId, mcid: node.number, form };
  }
  const object = value(); whitespace();
  assert.ok(!lookahead && position === body.length, "Unconsumed structure object tokens");
  assert.equal(object.kind, "dictionary");
  const page = object.fields.get("/Pg"), defaultPage = page === undefined ? undefined : reference(page, "Default page");
  const children = object.fields.get("/K"); assert.equal(children?.kind, "array", "Structure K must be an array");
  const entries = children.values.map(child => {
    if (child.kind === "reference") return { kind: "element", id: reference(child, "Structure child") };
    if (child.kind === "number") return mark(child, defaultPage, "integer");
    assert.equal(child.kind, "dictionary", "Unsupported structure child");
    const type = child.fields.get("/Type")?.text;
    assert.ok(type === "/MCR" || type === "/OBJR", "Unsupported structure child dictionary");
    const allowed = type === "/MCR" ? ["/Type", "/Pg", "/MCID"] : ["/Type", "/Pg", "/Obj"];
    assert.ok([...child.fields.keys()].every(key => allowed.includes(key)), "Unconsumed structure child dictionary fields");
    const ownPage = child.fields.get("/Pg"), pageId = ownPage === undefined ? defaultPage : reference(ownPage, "Content page");
    if (type === "/MCR") return mark(child.fields.get("/MCID"), pageId, "mcr");
    assert.ok(Number.isSafeInteger(pageId) && pageId > 0, "An annotation child requires a default or explicit page");
    return { kind: "annotation", id: reference(child.fields.get("/Obj"), "Annotation"), pageId };
  });
  return { defaultPage, kids: body.slice(children.start + 1, children.end - 1), entries };
}

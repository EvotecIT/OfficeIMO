import test from "node:test";
import assert from "node:assert/strict";
import { pdfStructureChildren } from "./pdf-structure-reader.mjs";

test("structure children preserve explicit MCRs, object references and annotation identity", () => {
  const body = "<< /Type /StructElem /S /P /P 1 0 R /K [12 0 R << /Type /MCR /Pg 3 0 R /MCID 0 >> << /Type /OBJR /Obj 9 0 R /Pg 3 0 R >>] >>";
  assert.deepEqual(pdfStructureChildren(body), { defaultPage: undefined,
    kids: "12 0 R << /Type /MCR /Pg 3 0 R /MCID 0 >> << /Type /OBJR /Obj 9 0 R /Pg 3 0 R >>",
    entries: [{ kind: "element", id: 12 }, { kind: "mark", pageId: 3, mcid: 0, form: "mcr" }, { kind: "annotation", id: 9, pageId: 3 }] });
});

test("integer marks use the outer page without inheriting a later MCR or OBJR override", () => {
  const parsed = pdfStructureChildren("<< /S /Link /K [0 1 << /Type /MCR /MCID 1 /Pg 9 0 R >> 2 << /Type /OBJR /Pg 9 0 R /Obj 8 0 R >> 15 0 R 3] /Pg 5 0 R /Alt (nested \\(text\\) /Pg 99 0 R) >>");
  assert.equal(parsed.defaultPage, 5);
  assert.deepEqual(parsed.entries, [
    { kind: "mark", pageId: 5, mcid: 0, form: "integer" }, { kind: "mark", pageId: 5, mcid: 1, form: "integer" },
    { kind: "mark", pageId: 9, mcid: 1, form: "mcr" },
    { kind: "mark", pageId: 5, mcid: 2, form: "integer" }, { kind: "annotation", id: 8, pageId: 9 },
    { kind: "element", id: 15 }, { kind: "mark", pageId: 5, mcid: 3, form: "integer" }
  ]);
});

test("OBJR-only continuation owners need no default page or marked text", () => {
  assert.deepEqual(pdfStructureChildren("<< /S /Link /K [<< /Type /OBJR /Obj 8 0 R /Pg 4 0 R >>] >>").entries,
    [{ kind: "annotation", id: 8, pageId: 4 }]);
  assert.deepEqual(pdfStructureChildren("<< /S /TD /Pg 4 0 R /K [<< /Type /MCR /MCID 0 >>] >>").entries,
    [{ kind: "mark", pageId: 4, mcid: 0, form: "mcr" }]);
});

test("invalid or partially consumed structure children fail instead of disappearing from semantic proof", () => {
  for (const kids of ["-1", "1.5", "9007199254740992", "0 0 R", "1 1 R", "0 R", "0 /Unexpected", "<< /Type /MCR /MCID 0 /Extra 1 >>", "<< /Type /MCR >>", "<< /Type /MCR /Pg 0 0 R /MCID 0 >>", "<< /Type /OBJR /Obj 0 0 R >>"])
    assert.throws(() => pdfStructureChildren("<< /S /P /Pg 3 0 R /K [" + kids + "] >>"), undefined, kids);
  assert.throws(() => pdfStructureChildren("<< /S /P /K [<< /Type /MCR /Pg 3 0 R /MCID 0 >> 1] >>"), /requires a default or explicit page/);
  assert.throws(() => pdfStructureChildren("<< /S /P /K [0] /A << /Pg 3 0 R >> >>"), /requires a default or explicit page/);
  assert.throws(() => pdfStructureChildren("<< /S /P /Pg 3 0 R /K [0] >> junk"), /Unconsumed/);
});

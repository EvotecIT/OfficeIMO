import test from "node:test";
import assert from "node:assert/strict";
import { writePdf, writePdfTo, writePdfDocument, pdfTable, ExportCell } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

// This fixture uses the standard font. General Unicode and consumer reading order are
// independently inspected by the real-report qualification runner.
async function logicalPdf(blob) {
  const pdf = await inspectPdf(blob), elements = new Map();
  for (const [id, object] of pdf.objects) if (/\/Type \/StructElem /.test(object.body)) {
    const kids = /\/K \[([\s\S]*?)\](?: |$)/.exec(object.body)[1], content = [];
    for (const token of kids.replace(/<< \/Type \/OBJR[^>]*>>/g, "").matchAll(/<< \/Type \/MCR \/Pg (\d+) 0 R \/MCID (\d+) >>|(\d+) 0 R/g)) {
      if (token[3]) content.push(Number(token[3]));
      else {
        const page = pdf.pages.find(value => value.id === Number(token[1]));
        const mark = new RegExp("/\\w+ << /MCID " + token[2] + "(?: /ActualText <[^>]*>)? >> BDC\\n([\\s\\S]*?)EMC").exec(page.content);
        assert.ok(mark);
        content.push([...mark[1].matchAll(/<([a-f0-9]*)> Tj/g)].map(value => Buffer.from(value[1], "hex").toString("latin1")).join(""));
      }
    }
    elements.set(id, { id, body: object.body, role: /\/S \/(\w+)/.exec(object.body)[1], content });
  }
  const text = id => elements.get(id).content.map(value => typeof value === "number" ? text(value) : value).join("");
  return { ...pdf, elements, text };
}

const columns = ["ID", "Detail", "Amount", "Count"].map(header => ({ header }));
const options = { columns, tagged: { language: "en" }, compression: false, pageNumbers: false,
  pageSize: { width: 300, height: 220 }, margins: 20, padding: 3, columnWidths: [60, 90, 90, 90],
  wideTable: { mode: "bands", repeatColumns: [0] }, limits: { maxTableWindowRows: 2 } };

test("tagged bands read each source cell once in row order across repeated identifiers and vertical fragments", async () => {
  let opened = 0, closed = 0, formatted = 0;
  const detail = Array.from({ length: 40 }, (_, index) => "LINE-" + index).join("\n");
  function* source() {
    opened++;
    try {
      yield [new ExportCell("ROW-1", { link: { target: "https://example.com/record/1" } }), detail, 12, 3];
      yield ["ROW-2", "Last", 24, 6];
    } finally { closed++; }
  }
  const pdf = await logicalPdf(await writePdf(source(), { ...options,
    formatValue(value) { formatted++; return String(value); },
    headerRows: [[{ value: "Identifier", rowSpan: 2 }, { value: "Metrics", columnSpan: 3 }, null, null],
      [null, { value: "Detail" }, { value: "Amount" }, { value: "Count" }]],
    footer: { rows: [[{ value: "Totals", columnSpan: 2 }, null, { value: 36 }, { value: 9 }]] }
  }));
  assert.equal(opened, 1); assert.equal(closed, 1); assert.equal(formatted, 8);
  assert.ok(pdf.pages.length > 4);
  const cells = [...pdf.elements.values()].filter(element => element.role === "TD");
  assert.deepEqual(cells.slice(0, 4).map(cell => pdf.text(cell.id)), ["ROW-1", detail.replaceAll("\n", ""), "12", "3"]);
  assert.deepEqual(cells.slice(4).map(cell => pdf.text(cell.id)), ["ROW-2", "Last", "24", "6", "Totals", "36", "9"]);
  const headers = [...pdf.elements.values()].filter(element => element.role === "TH");
  assert.deepEqual(headers.map(cell => pdf.text(cell.id)), ["Identifier", "Metrics", "Detail", "Amount", "Count"]);
  const links = [...pdf.elements.values()].filter(element => element.role === "Link");
  assert.equal(links.length, 1);
  assert.equal(pdf.text(links[0].id), "ROW-1");
  assert.ok((links[0].body.match(/\/Type \/OBJR/g) ?? []).length > 2, "Repeated visual links stay clickable with one logical link owner");
  assert.ok(pdf.pages.filter(page => page.lines.includes("ROW-1")).length > 2, "Identifiers repeat visually without duplicate semantic text");
});

test("tagged table captions precede logical headers and repeated continuation context stays artifact", async () => {
  const pdf = await logicalPdf(await writePdfDocument([
    { ...pdfTable([["R1", "detail", 10, 2], ["R2", "detail", 20, 4]], { columns, wideTable: options.wideTable, columnWidths: options.columnWidths }),
      caption: { text: "Table purpose" }, continuationCaption: { text: "Continued context" } }
  ], { ...options, columns: undefined }));
  const table = [...pdf.elements.values()].find(element => element.role === "Table");
  assert.equal(pdf.elements.get(table.content[0]).role, "Caption");
  assert.equal(pdf.text(table.content[0]), "Table purpose");
  assert.ok(![...pdf.elements.values()].some(element => pdf.text(element.id).includes("Continued context")));
});

test("tagged band structure limits return the source and release its failed destination", async () => {
  let returned = 0;
  const source = { async *[Symbol.asyncIterator]() { try { yield ["R1", "detail", 1, 2]; yield ["R2", "detail", 3, 4]; } finally { returned++; } } };
  const sink = new WritableStream({ write() {} });
  await assert.rejects(writePdfTo(source, sink, { ...options, limits: { maxStructureElements: 8 } }), /maxStructureElements/);
  assert.equal(returned, 1); assert.equal(sink.locked, false);
});

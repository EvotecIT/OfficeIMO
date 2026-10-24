import test from "node:test";
import assert from "node:assert/strict";
import { Workbook, Cell, StyleRegistry, NumberFormats } from "../dist/xlsx/index.js";
import { relationshipTypes } from "../dist/opc/index.js";
import { readZip } from "./zip-reader.mjs";

test("registered fonts/fills/borders/formats and custom column writers produce usable styles", async () => {
  const book = new Workbook({ cellValueWriters: { milliseconds: value => Number(value) / 1000 } });
  const font = book.styles.addFont({ bold: true, italic: true, color: "#123456" });
  assert.equal(font, book.styles.addFont({ bold: true, italic: true, color: "FF123456" }));
  const style = book.styles.add({ font, fill: { color: "ABCDEF" }, border: { bottom: { style: "thin", color: "123456" } }, numberFormat: NumberFormats.Decimal });
  const cellStyle = book.styles.add({ font: { italic: true }, numberFormat: "0.000" });
  await book.addWorksheet("Styles", { columns: [{ header: "Seconds", type: "milliseconds", style }] }).addRows([["1250"], [new Cell(2500, cellStyle)]]);
  book.addPart({ uri: "/customXml/item1.xml", contentType: "application/xml", data: '<data xmlns="urn:officeimo:test">Łódź</data>', relationship: { id: "custom", type: relationshipTypes.customXml } });
  const zip = await readZip(await book.toBlob());
  assert.match(zip.get("xl/styles.xml").content, /<i\/>/);
  assert.match(zip.get("xl/styles.xml").content, /bottom style="thin"/);
  assert.match(zip.get("xl/worksheets/sheet1.xml").content, /<v>1.25<\/v>/);
  assert.match(zip.get("xl/worksheets/sheet1.xml").content, new RegExp('r="A3" s="' + cellStyle + '"'));
  assert.match(zip.get("xl/_rels/workbook.xml.rels").content, /Target="..\/customXml\/item1.xml"/);
  assert.equal(book.worksheets[0].rowCount, 2);
});

test("workbook extension parts capture structural getters and preserve lazy byte producers", async () => {
  const book = new Workbook({ compression: "store" });
  const counts = { uri: 0, contentType: 0, data: 0, relationship: 0, id: 0, type: 0, source: 0 };
  let produced = 0;
  const source = async sink => { produced++; await sink.write(Buffer.from("<custom/>")); };
  class Link {
    get id() { counts.id++; return "custom"; }
    get type() { counts.type++; return relationshipTypes.customXml; }
    get source() { counts.source++; return "/xl/workbook.xml"; }
  }
  class Part {
    get uri() { counts.uri++; return "/customXml/item.xml"; }
    get contentType() { counts.contentType++; return "application/xml"; }
    get data() { counts.data++; return source; }
    get relationship() { counts.relationship++; return new Link(); }
  }
  book.addPart(new Part());
  assert.equal(produced, 0);
  const zip = await readZip(await book.toBlob());
  assert.equal(zip.get("customXml/item.xml").content, "<custom/>");
  assert.ok(zip.get("xl/_rels/workbook.xml.rels").content.includes(
    'Id="custom" Type="' + relationshipTypes.customXml + '" Target="../customXml/item.xml"'));
  assert.equal(produced, 1);
  assert.deepEqual(counts, { uri: 1, contentType: 1, data: 1, relationship: 1, id: 1, type: 1, source: 1 });
});

test("XML reject policy reaches worksheet text and registered styles", async () => {
  const book = new Workbook({ invalidCharacterPolicy: "reject" });
  assert.throws(() => new Workbook({ invalidCharacterPolicy: "reject", creator: "bad\u0001" }), { code: "INVALID_XML" });
  assert.throws(() => book.styles.addFont({ name: "bad\u0001" }), { code: "INVALID_XML" });
  assert.throws(() => book.styles.addNumberFormat("0\u0001"), { code: "INVALID_XML" });
  await assert.rejects(book.addWorksheet("Data", { columns: [{ header: "V" }] }).addRows([["bad\u0001"]]), { code: "INVALID_XML" });
  await assert.rejects(book.toBlob(), { code: "INVALID_XML" });
  assert.throws(() => new StyleRegistry().add({ font: 1 }), RangeError);
});

test("workbook extensions reject generated part collisions before export begins", async () => {
  for (const uri of ["/XL", "/xl/workbook.xml", "/XL/WORKBOOK.XML/extension.xml", "/xl/styles.xml/extension.xml"])
    assert.throws(() => new Workbook().addPart({ uri, contentType: "application/xml", data: "<extension/>" }), /Duplicate|prefix collision/i);
  const book = new Workbook();
  book.addWorksheet("Data");
  book.addPart({ uri: "/xl/worksheets/sheet1.xml/extension.xml", contentType: "application/xml", data: "<extension/>" });
  await assert.rejects(book.toBlob(), /prefix collision/i);
});

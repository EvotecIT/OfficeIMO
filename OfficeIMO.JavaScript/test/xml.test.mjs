import test from "node:test";
import assert from "node:assert/strict";
import { XmlWriter, cleanXml, escapeXml } from "../dist/xml/index.js";
import { BlobByteSink } from "../dist/core/index.js";

test("XML writer escapes data, normalizes no whitespace and validates schema-owned names", async () => {
  const sink = new BlobByteSink(), writer = new XmlWriter(sink);
  await writer.startElement("root", { "xmlns": "urn:test", "value": '"<&\r\n\t' });
  await writer.startElement("child"); await writer.text("Łódź 🧪 שלום\r\n\t<&\u0001");
  await writer.endElement(); await writer.endElement(); await writer.close();
  const text = await sink.toBlob().text();
  assert.match(text, /value="&quot;&lt;&amp;&#13;&#10;&#9;"/);
  assert.match(text, /Łódź 🧪 שלום&#13;&#10;&#9;&lt;&amp;<\/child>/);
  await assert.rejects(writer.text("late"), { code: "INVALID_STATE" });
  const names = new XmlWriter(new BlobByteSink());
  await assert.rejects(names.startElement('x><evil'), { code: "INVALID_XML" });
  await assert.rejects(names.startElement("root", { 'x" bad': "data" }), { code: "INVALID_XML" });
  await names.startElement("root"); await names.endElement(); await names.close();
});

test("XML reject policy reports invalid controls and incomplete documents", async () => {
  assert.equal(cleanXml("🧪\ud800\u0000ok"), "🧪ok");
  assert.throws(() => escapeXml("\u0001", "reject"), { code: "INVALID_XML" });
  const writer = new XmlWriter(new BlobByteSink(), "reject");
  await writer.startElement("r"); await assert.rejects(writer.text("\ud800"), { code: "INVALID_XML" });
  await assert.rejects(writer.close(), { code: "INVALID_XML" });
});

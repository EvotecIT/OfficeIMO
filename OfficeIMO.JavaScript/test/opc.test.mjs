import test from "node:test";
import assert from "node:assert/strict";
import { ContentTypes, OpcPackage, partUri, relativePartTarget, relationshipPartUri, relationshipTypes } from "../dist/opc/index.js";
import { readZip } from "./zip-reader.mjs";

test("OPC owns content types, relative relationships and core/app property documents", async () => {
  const packageFile = new OpcPackage({ compression: "store" });
  packageFile.addPart({ uri: "/word/document.xml", contentType: "application/xml", data: "<document/>" });
  packageFile.addPart({ uri: "/customXml/item1.xml", contentType: "application/xml", data: "<custom/>" });
  const bytes = Buffer.from("<original/>");
  packageFile.addPart({ uri: "/bytes.xml", contentType: "application/xml", data: bytes });
  bytes.fill(120);
  packageFile.addRelationship("/", { id: "document", type: relationshipTypes.officeDocument, target: "/word/document.xml" });
  packageFile.addRelationship("/word/document.xml", { id: "custom", type: relationshipTypes.customXml, target: "/customXml/item1.xml" });
  packageFile.setProperties({ creator: "A<&", created: new Date("2026-10-05T00:00:00Z") }, { company: "Evotec" });
  const zip = await readZip(await packageFile.toBlob());
  assert.match(zip.get("word/_rels/document.xml.rels").content, /Target="..\/customXml\/item1.xml"/);
  assert.match(zip.get("[Content_Types].xml").content, /PartName="\/word\/document.xml"/);
  assert.match(zip.get("docProps/core.xml").content, /A&lt;&amp;/);
  assert.match(zip.get("docProps/app.xml").content, /Evotec/);
  assert.equal(zip.get("bytes.xml").content, "<original/>");
  assert.equal(relationshipPartUri("/"), "/_rels/.rels");
  assert.equal(relativePartTarget("/a/b.xml", "/a/c.xml"), "c.xml");
  assert.throws(() => packageFile.addPart({ uri: "/late.xml", contentType: "application/xml", data: "" }), { code: "INVALID_STATE" });
});

test("OPC rejects malformed/colliding part URIs and unresolved relationships", async () => {
  for (const uri of ["relative", "/", "/a/../b", "/a%2Fb", "/%61.xml", "/bad%", "/bad.", "/a?x", "/a#x", "/Łódź.xml", '/quote".xml', "/angle<.xml", "/bracket[.xml"])
    assert.throws(() => partUri(uri), { code: "INVALID_PART_URI" });
  assert.equal(partUri("/caf%c3%a9.xml"), "/caf%C3%A9.xml");
  const packageFile = new OpcPackage();
  packageFile.addPart({ uri: "/a.xml", contentType: "application/xml", data: "<a/>" });
  assert.throws(() => packageFile.addPart({ uri: "/A.xml", contentType: "application/xml", data: "<b/>" }), /Duplicate/);
  assert.throws(() => packageFile.addPart({ uri: "/_rels/.rels", contentType: "application/xml", data: "" }), /reserved/);
  packageFile.addRelationship("/", { id: "missing", type: relationshipTypes.officeDocument, target: "/missing.xml" });
  await assert.rejects(packageFile.toBlob(), /Missing relationship target/);
  const types = new ContentTypes(); types.addDefault("png", "image/png");
  assert.throws(() => types.addDefault("png", "image/jpeg"), /Conflicting/);
});

test("unreadable Blob parts fail with a typed host error and preserve the native cause", async () => {
  class UnreadableBlob extends Blob {
    slice() { return this; }
    arrayBuffer() { return Promise.reject(new DOMException("Worker Blob reads are blocked.", "NotReadableError")); }
  }
  const packageFile = new OpcPackage();
  packageFile.addPart({ uri: "/data.xml", contentType: "application/xml", data: new UnreadableBlob(["<data/>"]) });
  await assert.rejects(packageFile.toBlob(), error => error.code === "PLATFORM_UNAVAILABLE" && error.cause?.name === "NotReadableError" && /byte chunks/.test(error.message));
  await assert.rejects(packageFile.toBlob(), { code: "PLATFORM_UNAVAILABLE" });
});

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

test("OPC captures structural part and relationship fields once before retaining them", async () => {
  for (const ownFields of [false, true]) {
    const packageFile = new OpcPackage({ compression: "store" });
    const counts = { uri: 0, contentType: 0, data: 0, id: 0, type: 0, target: 0, external: 0 };
    const bytes = Buffer.from("<first/>");
    class Part {
      get uri() { counts.uri++; return "/custom.xml"; }
      get contentType() { counts.contentType++; return "application/xml"; }
      get data() { return ++counts.data === 1 ? bytes : Buffer.from("<later/>"); }
    }
    class Link {
      get id() { counts.id++; return "custom"; }
      get type() { counts.type++; return relationshipTypes.customXml; }
      get target() { counts.target++; return ownFields ? "https://example.com/report" : "/custom.xml"; }
      get external() { counts.external++; return ownFields; }
    }
    const record = Type => ownFields ? Object.create(null, Object.getOwnPropertyDescriptors(Type.prototype)) : new Type();
    packageFile.addPart(record(Part));
    packageFile.addRelationship("/", record(Link));
    bytes.fill(120);
    const zip = await readZip(await packageFile.toBlob());
    assert.equal(zip.get("custom.xml").content, "<first/>");
    const relationships = zip.get("_rels/.rels").content;
    assert.ok(relationships.includes('Id="custom" Type="' + relationshipTypes.customXml + '"'));
    assert.ok(relationships.includes(ownFields ? 'Target="https://example.com/report" TargetMode="External"' : 'Target="custom.xml"'));
    assert.deepEqual(counts, { uri: 1, contentType: 1, data: 1, id: 1, type: 1, target: 1, external: 1 });
  }
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

test("OPC rejects part ancestors and descendants case-insensitively in either registration order", async () => {
  for (const uris of [["/Data", "/data/child.xml"], ["/DATA/child.xml", "/data"]]) {
    const packageFile = new OpcPackage({ compression: "store" });
    packageFile.addPart({ uri: uris[0], contentType: "application/xml", data: "<first/>" });
    assert.throws(() => packageFile.addPart({ uri: uris[1], contentType: "application/xml", data: "<second/>" }), /prefix collision/i);
    // A rejected name must not reserve unrelated ancestors or break the existing package.
    packageFile.addPart({ uri: "/database/child.xml", contentType: "application/xml", data: "<sibling/>" });
    assert.equal((await readZip(await packageFile.toBlob())).size, 3);
  }
});

test("OPC protects generated property and relationship metadata paths", () => {
  const properties = new OpcPackage(); properties.setProperties();
  for (const uri of ["/DOCPROPS", "/docProps/core.xml/child.xml", "/docProps/app.xml/child.xml"])
    assert.throws(() => properties.addPart({ uri, contentType: "application/xml", data: "<data/>" }), /prefix collision/i);
  const reverse = new OpcPackage(); reverse.addPart({ uri: "/docProps", contentType: "application/xml", data: "<data/>" });
  assert.throws(() => reverse.setProperties(), /prefix collision/i);
  for (const uri of ["/_RELS", "/word/_rels", "/word/_RELS/document.xml.rels/child.xml"])
    assert.throws(() => new OpcPackage().addPart({ uri, contentType: "application/xml", data: "<data/>" }), /reserved/i);
  for (const uri of ["/[Content_Types].xml", "/[Content_Types].xml/child.xml"])
    assert.throws(() => new OpcPackage().addPart({ uri, contentType: "application/xml", data: "<data/>" }));
});

test("relative relationship targets with a colon in the first segment remain internal URIs", async () => {
  for (const [source, target, expected] of [["/", "/custom:part.xml", "./custom:part.xml"],
    ["/word/document.xml", "/word/custom:part.xml", "./custom:part.xml"],
    ["/word/document.xml", "/custom:part.xml", "../custom:part.xml"],
    ["/", "/word/custom:part.xml", "word/custom:part.xml"]]) {
    assert.equal(relativePartTarget(source, target), expected);
    const base = new URL(source, "https://officeimo.invalid/");
    assert.equal(new URL(expected, base).href, "https://officeimo.invalid" + target);
    const packageFile = new OpcPackage({ compression: "store" });
    if (source !== "/") packageFile.addPart({ uri: source, contentType: "application/xml", data: "<document/>" });
    packageFile.addPart({ uri: target, contentType: "application/xml", data: "<custom/>" });
    packageFile.addRelationship(source, { id: "custom", type: relationshipTypes.customXml, target });
    const zip = await readZip(await packageFile.toBlob());
    assert.ok(zip.get(relationshipPartUri(source).slice(1)).content.includes('Target="' + expected + '"'));
  }
});

test("relationship IDs are unique per source and retain insertion order across streamed XML chunks", async () => {
  const packageFile = new OpcPackage({ compression: "store" });
  for (const uri of ["/first.xml", "/second.xml"])
    packageFile.addPart({ uri, contentType: "application/xml", data: "<document/>" });
  const relationship = i => ({ id: "link" + i, type: relationshipTypes.officeDocument, target: "https://example.com/?id=" + i + "&mode=report", external: true });
  for (let i = 0; i < 2048; i++) packageFile.addRelationship("/first.xml", relationship(i));
  assert.throws(() => packageFile.addRelationship("/FIRST.xml", relationship(1024)), /Duplicate relationship id/);
  packageFile.addRelationship("/second.xml", relationship(1024));
  const zip = await readZip(await packageFile.toBlob());
  const first = zip.get("_rels/first.xml.rels").content;
  const ids = [...first.matchAll(/<Relationship Id="([^"]+)"/g)].map(match => match[1]);
  assert.deepEqual(ids, Array.from({ length: 2048 }, (_, i) => "link" + i));
  assert.match(first, /id=2047&amp;mode=report/);
  assert.equal([...zip.get("_rels/second.xml.rels").content.matchAll(/<Relationship /g)].length, 1);
});

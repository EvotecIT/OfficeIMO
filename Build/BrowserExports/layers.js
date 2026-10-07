// Browser assertions exercise the public graph and the shared corpus, not a browser-specific writer.
async function runLayerScenarios({ fixtureJson, moduleBase }) {
  const libraries = [["classic", OfficeIMO]], fixtures = JSON.parse(fixtureJson);
  let assertions = 0;
  const require = (condition, message) => { assertions++; if (!condition) throw new Error(message); };
  if (moduleBase) {
    const library = await import(moduleBase + "/index.js");
    for (const layer of ["core", "zip", "xml", "opc", "xlsx", "csv"]) {
      const module = await import(moduleBase + "/" + layer + "/index.js");
      for (const [key, value] of Object.entries(module)) require(library[layer][key] === value, "ESM namespace differs: " + layer + "/" + key);
    }
    libraries.push(["esm", library]);
  }
  for (const [kind, library] of libraries) {
    const { core, zip, xml, opc, xlsx, csv } = library;
    const sink = new core.BlobByteSink(), writer = new xml.XmlWriter(sink);
    await writer.startElement("root", { value: "<&\r\n\t" }); await writer.text("Łódź 🧪 שלום"); await writer.endElement(); await writer.close();
    const parsed = new DOMParser().parseFromString(await sink.toBlob().text(), "application/xml");
    require(parsed.documentElement.textContent === "Łódź 🧪 שלום" && parsed.documentElement.getAttribute("value") === "<&\r\n\t", "XML parser round trip differs");
    const archive = new zip.ZipWriter(); await archive.add("test.txt", new TextEncoder().encode("Łódź"));
    require((await archive.toBlob()).size > 0, "ZIP layer failed");
    const packageFile = new opc.OpcPackage(); packageFile.addPart({ uri: "/test.xml", contentType: "application/xml", data: "<test/>" });
    packageFile.addRelationship("/", { id: "test", type: opc.relationshipTypes.officeDocument, target: "/test.xml" });
    require((await packageFile.toBlob()).size > 0, "OPC layer failed");
    require(await (await csv.writeCsv([["=cmd"]], { columns: [{ header: "V" }] })).text() === "V\r\n'=cmd\r\n", "CSV layer failed");
    require(await (await csv.writeCsv([["Łódź", false]], { columns: [{ header: "Name", valueFormatter: value => "=" + value },
      { header: "Healthy", valueFormatter: value => value ? "Yes" : "No" }], quote: "all" })).text() === '"Name","Healthy"\r\n"\'=Łódź","No"\r\n', "Formatted CSV protection/quoting differs");
    const frame = document.createElement("iframe"), loaded = new Promise(resolve => { frame.onload = resolve; });
    frame.srcdoc = "<!doctype html><title>Stream destination realm</title>"; document.body.append(frame); await loaded;
    try {
      const ForeignStream = frame.contentWindow.WritableStream;
      for (const [format, write] of [["csv", csv.writeCsvTo], ["xlsx", xlsx.writeXlsxTo]]) {
        const chunks = []; let closes = 0, aborts = 0;
        const destination = new ForeignStream({ write: bytes => { chunks.push(new Uint8Array(bytes)); },
          close: () => { closes++; }, abort: () => { aborts++; } });
        require(!(destination instanceof WritableStream), "Destination must exercise another realm");
        const result = await write([["Łódź 🧪", 12.5]], destination, { columns: [{ header: "Name" }, { header: "Amount", type: "number" }] });
        const blob = new Blob(chunks);
        require(!destination.locked && closes === 0 && aborts === 0 && result.rows === 1 && result.columns === 2 && result.bytes === blob.size, "Foreign " + format + " stream ownership/result differs");
        if (format === "csv") require(await blob.text() === "Name,Amount\r\nŁódź 🧪,12.5\r\n", "Foreign CSV bytes differ");
        else await emitFixture("realm-" + kind + ".xlsx", blob);
        const failure = new Error("Foreign destination failed"), failing = new ForeignStream({ write() { throw failure; } });
        let rejected; try { await write([[1]], failing, { columns: [{ header: "Value" }] }); } catch (error) { rejected = error; }
        require(rejected === failure && !failing.locked, "Foreign " + format + " rejection/lock release differs");
      }
    } finally { frame.remove(); }
    for (const spec of fixtures.cases) for (const compression of ["auto", "store"]) {
      if (spec.producer === "table-helper") {
        const sheet = spec.sheets[0], chunks = [];
        const columns = sheet.columns.map(column => column.key === "name" ? { ...column, value: row => row.person.name } : column);
        const rows = sheet.rows.map(row => ({ person: { name: row[0] }, amount: new core.ExportCell(row[1], { presentation: { background: "C6EFCE" } }), seen: new Date(row[2].value), ignored: { domain: true } }));
        const options = { columns, compression, dateMode: "utc", sheet };
        let blob;
        if (compression === "auto") blob = await xlsx.writeXlsx(rows, options);
        else {
          const stream = new WritableStream({ write: bytes => { chunks.push(new Uint8Array(bytes)); } });
          const result = await xlsx.writeXlsxTo(rows, stream, options); blob = new Blob(chunks);
          require(!stream.locked && result.rows === rows.length && result.columns === columns.length && result.bytes === blob.size, "Table helper result/stream ownership differs");
        }
        require(await (await csv.writeCsv(rows, { columns })).text() === "Name,Amount,Seen\r\nŁódź 🧪,12.5,2026-10-07T00:00:00.000Z\r\nWarsaw,125.75,2026-10-08T00:00:00.000Z\r\n", "Shared domain projection differs in CSV");
        await emitFixture("corpus-" + kind + "-" + spec.name + "-" + compression + ".xlsx", blob);
        continue;
      }
      const book = new xlsx.Workbook({ dateMode: "utc", compression, created: new Date("2026-10-05T00:00:00Z"),
        cellValueWriters: { milliseconds: value => Number(value) / 1000 }, ...spec.options });
      const styles = (spec.styles ?? []).map(style => book.styles.add(style));
      const value = v => v?.kind === "date" ? new Date(v.value) : v?.kind === "cell" ? new xlsx.Cell(value(v.value), styles[v.style]) : v;
      for (const sheet of spec.sheets) {
        const columns = sheet.columns.map(c => c.style === undefined ? c : { ...c, style: styles[c.style] });
        const worksheet = book.addWorksheet(sheet.name, { ...sheet, columns,
          ...(sheet.headerStyle === undefined ? {} : { headerStyle: styles[sheet.headerStyle] }),
          ...(sheet.statusHighlight ? { rowStyle: ({ values }) => values[3] === false ? { fill: { color: "FCE4D6" }, font: { bold: true } } : undefined,
            cellStyle: ({ value, columnIndex }) => columnIndex === 1 && value > 100 ? { font: { color: "C00000" } } : undefined } : {}) });
        await worksheet.addRows(sheet.rows.map(row => Array.isArray(row) ? row.map(value) :
          Object.fromEntries(Object.entries(row).map(([key, cell]) => [key, value(cell)]))));
        for (const image of sheet.images ?? []) worksheet.addImage({ ...image, data: Uint8Array.from(atob(image.pngBase64), ch => ch.charCodeAt(0)) });
      }
      for (const part of spec.parts ?? []) book.addPart(part);
      await emitFixture("corpus-" + kind + "-" + spec.name + "-" + compression + ".xlsx", await book.toBlob());
    }
    for (const feature of ["dataValidation"]) {
      try { new xlsx.Workbook().addWorksheet("Data", { [feature]: [] }); throw new Error("Ignored reserved option"); }
      catch (error) { require(error.code === "NOT_SUPPORTED" && error.feature === feature, "Reserved option did not throw clearly"); }
    }
  }
  return { assertions, classicSubpaths: 6, esmSubpaths: moduleBase ? 6 : 0, workbooks: libraries.length * fixtures.cases.length * 2 };
}

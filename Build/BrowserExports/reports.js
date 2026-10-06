// Small final-contract lane complements the streamed scale matrix with native-reader artifacts.
async function runReportContracts(fixtureJson) {
  const chart = JSON.parse(fixtureJson).cases.find(spec => spec.name === "report-table").sheets[0].images[0];
  const operations = ["sum", "count", "average", "min", "max"];
  for (const compression of ["auto", "store"]) {
    for (const streamed of [false, true]) {
      const chunks = [], reports = OfficeIMO.createWorkbook({ compression, cellValueWriters: { milliseconds: value => Number(value) / 1000 }, ...(streamed ? { sink: { write(bytes) { chunks.push(bytes.slice()); } } } : {}) });
      const titled = reports.addSheet("Report", { title: { text: "Łódź 🧪 report", height: 32 }, columns: [{ header: "Name", groups: ["Metrics"] }, { header: "Amount", key: "amount", groups: ["Metrics"], type: "number" }], table: { name: "Report" }, freezeHeader: true, autoSize: {}, footer: { values: ["Total"], totals: { amount: "sum" } }, print: { repeatHeaders: true } });
      await titled.addRows([["one", 12], ["two", 8]]); await titled.close();
      const regions = reports.addSheet("Regions", { columns: [{ header: "A" }, { header: "B", type: "milliseconds" }], includeHeader: false, mergedCells: ["A1:B2", "A3:B3"], autoSize: { sampleRows: 3 } });
      await regions.addRows([["top", ""], [null, new OfficeIMO.ExportCell(null)], ["bottom", new OfficeIMO.Cell("")]]); await regions.close();
      const third = reports.addSheet("Other report", { title: { text: "Other report" }, columns: [{ header: "Metric" }, { header: "Value", type: "number" }], table: { name: "OtherReport" } });
      await third.addRows([["Count", 2]]);
      await third.close();
      const charts = reports.addSheet("Charts", { title: { text: "Report latency" }, columns: Array.from({ length: 6 }, (_, i) => ({ header: "Column " + i, width: 12 })), includeHeader: false, print: {} });
      await charts.addRows(Array.from({ length: 18 }, () => Array(6).fill(null)));
      charts.addImage({ data: Uint8Array.from(atob(chart.pngBase64), ch => ch.charCodeAt(0)), row: 3, column: 1, width: chart.width, height: chart.height, description: "ChartForgeX report latency" });
      let blob; if (streamed) { await reports.finish(); blob = new Blob(chunks); } else blob = await reports.toBlob();
      await emit("report-regions-" + compression + "-" + (streamed ? "stream" : "blob") + ".xlsx", blob);
      const invalid = OfficeIMO.createWorkbook({ compression });
      const hidden = invalid.addSheet("Hidden", { columns: [{ header: "A" }, { header: "B" }], includeHeader: false, mergedCells: ["A1:B1"] });
      try { await hidden.addRows([["anchor", 0]]); throw new Error("Merged value was lost"); } catch (error) { if (!String(error).includes("would hide")) throw error; }
    }
    const dates = OfficeIMO.createWorkbook({ dateMode: "utc", compression });
    const sheet = dates.addSheet("Dates", { columns: operations.map(header => ({ header, type: "date", format: "yyyy-mm-dd" })), autoSize: {}, table: { name: "DateTotals" },
      footer: { totals: Object.fromEntries(operations.map(operation => [operation, operation])) } });
    await sheet.addRows([operations.map(() => new Date("2026-10-06T00:00:00Z")), operations.map(() => new OfficeIMO.ExportCell(new Date("2026-10-08T00:00:00Z")))]);
    await emit("report-dates-" + compression + ".xlsx", await dates.toBlob());
    const footer = OfficeIMO.createWorkbook({ compression, oversizedText: "preserve", limits: { maxTextCharacters: 140000 } });
    await footer.addSheet("Footer", { columns: [{ header: "A" }], footer: { values: ["x".repeat(100000)] } }).addRows([[1]]);
    await emit("report-preserved-footer-" + compression + ".xlsx", await footer.toBlob());
    for (const closed of [false, true]) {
      const controller = new AbortController(), reason = new Error("cancel between calls");
      const book = OfficeIMO.createWorkbook({ compression, signal: controller.signal });
      const cancelled = book.addSheet("Cancel", { columns: [{ header: "A" }] });
      await cancelled.addRows([["accepted"]]); if (closed) await cancelled.close(); controller.abort(reason);
      try { await book.finish(); throw new Error("Cancellation succeeded unexpectedly"); } catch (error) { if (error !== reason) throw error; }
    }
  }
  return { passed: true, dateArtifacts: 2, preservedFooterArtifacts: 2, preflightCancellationCases: 4, mergedReportArtifacts: 4, hiddenValueRejections: 4 };
  async function emit(name, blob) {
    const bytes = new Uint8Array(await blob.arrayBuffer());
    let binary = ""; for (let offset = 0; offset < bytes.length; offset += 8192) binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
    await writeFixture(name, btoa(binary));
  }
}

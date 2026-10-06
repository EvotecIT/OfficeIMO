// Small final-contract lane complements the streamed scale matrix with native-reader artifacts.
async function runReportContracts() {
  const operations = ["sum", "count", "average", "min", "max"];
  for (const compression of ["auto", "store"]) {
    const dates = OfficeIMO.createWorkbook({ dateMode: "utc", compression });
    const sheet = dates.addSheet("Dates", { columns: operations.map(header => ({ header, type: "date", format: "yyyy-mm-dd" })), autoSize: {},
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
  return { passed: true, dateArtifacts: 2, preservedFooterArtifacts: 2, preflightCancellationCases: 4 };
  async function emit(name, blob) {
    const bytes = new Uint8Array(await blob.arrayBuffer());
    let binary = ""; for (let offset = 0; offset < bytes.length; offset += 8192) binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
    await writeFixture(name, btoa(binary));
  }
}

/* Values, styles, continuation geometry and native link actions share one captured cell. */
async function makePortableLinkFixtures(streamed = false) {
  const { ExportCell, writeXlsx, writeXlsxTo, writePdf, writePdfTo, writeCsv, writeCsvTo } = OfficeIMO;
  async function bytes(writer, to, rows, options) {
    if (!streamed) return new Uint8Array(await (await writer(rows, options)).arrayBuffer());
    const chunks = [];
    await to(rows, { write(chunk) { chunks.push(chunk.slice()); } }, options);
    const result = new Uint8Array(chunks.reduce((n, chunk) => n + chunk.length, 0));
    let offset = 0; for (const chunk of chunks) { result.set(chunk, offset); offset += chunk.length; }
    return result;
  }
  const link = { target: 'https://example.com/report?q=1&other=2#details', tooltip: 'Report _x0041_ <value> 🧪' };
  const rows = [[new ExportCell(12.5, { text: '12.50 USD', link, presentation: { background: 'E2F0D9' } }),
    new ExportCell(new Date('2026-10-08T00:00:00Z'), { link: { target: 'mailto:report@example.com' } })]];
  const xlsx = await bytes(writeXlsx, writeXlsxTo, rows, { dateMode: 'utc', columns: [
    { header: 'Amount', type: 'number', format: '0.00', groups: ['Inventory'] },
    { header: 'Seen', type: 'date', groups: ['Inventory'] }
  ], sheet: { title: { text: 'Report' }, footer: { values: [new ExportCell('Footer', { link }), null] } } });
  const pdf = await bytes(writePdf, writePdfTo, [[new ExportCell('ABCDEFGHIJKLMNOPQRSTUVWXYZ'.repeat(300), { link })]], {
    columns: [{ header: 'Value' }], pageSize: 'A5', columnWidths: [160]
  });
  const spans = await bytes(writePdf, writePdfTo, Array.from({ length: 60 }, (_, i) => ['Row ' + i, i]), {
    columns: [{ header: 'Name' }, { header: 'Number' }], pageSize: 'A5',
    headerRows: [[{ value: new ExportCell('Report', { link }), columnSpan: 2 }, null]],
    footer: { rows: [[{ value: new ExportCell('End', { link: { target: 'mailto:report@example.com' } }), columnSpan: 2 }, null]] }
  });
  for (const valueMode of ['raw', 'display']) {
    const csv = await bytes(writeCsv, writeCsvTo, [[rows[0][0]]], { columns: [{ header: 'Amount' }], valueMode, includeHeader: false });
    if (new TextDecoder().decode(csv) !== (valueMode === 'raw' ? '12.5' : '12.50 USD') + '\r\n') throw new Error('Portable CSV value semantics differ.');
  }
  return [['portable-links.xlsx', xlsx], ['portable-links.pdf', pdf], ['portable-links-spans.pdf', spans]]
    .map(([name, bytes]) => ({ name, bytes }));
}

globalThis.runPortableLinkContracts = async function (workerScript) {
  const produced = [];
  for (const mode of ['classic', 'fallback', 'worker']) {
    let fixtures;
    if (mode === 'worker') {
      const source = workerScript + '\n' + makePortableLinkFixtures.toString() + '\n' +
        "onmessage = async () => { try { const files = await makePortableLinkFixtures(true); postMessage({ files }, files.map(f => f.bytes.buffer)); } catch (e) { postMessage({ error: e.stack || String(e) }); } };";
      const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
      const worker = new Worker(url);
      try {
        fixtures = await new Promise((resolve, reject) => {
          worker.onmessage = event => event.data.error ? reject(new Error(event.data.error)) : resolve(event.data.files);
          worker.onerror = event => reject(new Error(event.message)); worker.postMessage({});
        });
      } finally { worker.terminate(); URL.revokeObjectURL(url); }
    } else {
      const compressor = globalThis.CompressionStream;
      try { if (mode === 'fallback') globalThis.CompressionStream = undefined; fixtures = await makePortableLinkFixtures(); }
      finally { globalThis.CompressionStream = compressor; }
    }
    for (const file of fixtures) {
      let binary = '';
      for (let i = 0; i < file.bytes.length; i += 32768) binary += String.fromCharCode(...file.bytes.subarray(i, i + 32768));
      await writePortableLinkFixture(mode + '-' + file.name, btoa(binary)); produced.push(mode + '-' + file.name);
    }
  }
  return { files: produced, modes: ['classic', 'fallback', 'worker'], csvValueModes: ['raw', 'display'] };
};

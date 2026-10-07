// Qualification of a host-owned, demand-driven handoff. DataTables remains on the page.
globalThis.runDataTablesWorker = async function (source, script, compression, cancelAfterRows) {
  const bootstrap = `
    let sequence = 0, controller; const pending = new Map();
    const ask = (kind, payload) => new Promise((resolve, reject) => { const id = ++sequence; pending.set(id, {resolve,reject}); postMessage({id,kind,...payload}); });
    onmessage = async ({data}) => {
      if (data.cancel) { controller.abort('worker cancelled'); return; }
      if (data.id) { const request = pending.get(data.id); pending.delete(data.id); data.error ? request.reject(new Error(data.error)) : request.resolve(data); return; }
      controller = new AbortController();
      try {
        const sink = { async write(chunk) { for (let first = 0; first < chunk.length; first += 65536) {
          const bytes = chunk.slice(first, first + 65536); await ask('bytes', {bytes});
        } } };
        async function* rows() { while (true) { const batch = await ask('rows', {});
          for (const row of batch.rows) yield row.map(cell => cell.resolved ? new OfficeIMO.ExportCell(cell.value, cell.options) : cell.value);
          if (batch.done) return;
        } }
        const book = new OfficeIMO.Workbook({sink,signal:controller.signal,compression:data.compression});
        await book.addSheet('Worker',{columns:data.columns}).addRows(rows()); await book.finish(); postMessage({done:true});
      } catch (error) { postMessage({failure:String(error)}); }
    };`;
  const url = URL.createObjectURL(new Blob([script, '\n', bootstrap], { type: 'text/javascript' }));
  const worker = new Worker(url), iterator = source.rows[Symbol.asyncIterator](), chunks = [];
  let produced = 0, maxBatch = 0, maxChunk = 0, timeout;
  try {
    const result = await new Promise((resolve, reject) => {
      timeout = setTimeout(() => reject(new Error('Table worker timed out.')), 30000);
      worker.onerror = event => reject(new Error(event.message));
      worker.onmessage = async ({data}) => {
        try {
          if (data.kind === 'rows') {
            const rows = []; let done = false;
            while (rows.length < 64) {
              const next = await iterator.next(); if (next.done) { done = true; break; }
              rows.push(next.value.map(cell => cell instanceof OfficeIMO.ExportCell
                ? { resolved:true,value:cell.value,options:{text:cell.text,presentation:cell.presentation} } : {value:cell}));
            }
            produced += rows.length; maxBatch = Math.max(maxBatch, rows.length);
            if (cancelAfterRows && produced >= cancelAfterRows) worker.postMessage({cancel:true});
            worker.postMessage({id:data.id,rows,done});
          } else if (data.kind === 'bytes') {
            maxChunk = Math.max(maxChunk,data.bytes.byteLength); chunks.push(data.bytes);
            await new Promise(resolve => setTimeout(resolve,1)); worker.postMessage({id:data.id});
          } else if (data.failure) resolve({failure:data.failure});
          else if (data.done) resolve({blob:new Blob(chunks)});
        } catch (error) { reject(error); }
      };
      worker.postMessage({columns:source.columns,compression});
    });
    return {...result,produced,maxBatch,maxChunk};
  } finally { clearTimeout(timeout); worker.terminate(); URL.revokeObjectURL(url); await iterator.return?.(); }
};

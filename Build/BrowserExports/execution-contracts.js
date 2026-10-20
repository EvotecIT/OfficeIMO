// Product contracts use real host-owned Workers. HtmlTinkerX owns browser sessions.
async function runExecutionContracts({ moduleBase, classicScript, executionScript, fontBase64, classic = false }) {
  const api = classic ? globalThis.OfficeIMO : await import(moduleBase + '/index.js');
  const execution = classic ? api : await import(moduleBase + '/execution/index.js');
  const require = (value, message) => { if (!value) throw Error(message); };
  const checks = [], metrics = [];
  const fontBytes = Uint8Array.from(atob(fontBase64), character => character.charCodeAt(0));
  const columns = [{ header: 'Name' }, { header: 'Amount', type: 'number' }, { header: 'Flag', type: 'boolean' }];
  const workerBody = `
    const handler = execution.handleWorkerExports(self, async ({input,rows,sink,signal,onProgress}) => {
      if(input.mode === 'failure') throw new api.core.OfficeIMOError('RESOURCE_LIMIT','worker resource failure');
      if(input.mode === 'dispose') { setTimeout(() => handler.dispose(),0); return new Promise(() => {}); }
      if(input.mode === 'progress') { for(let i=0;i<1000;i++) onProgress({phase:'rows',rows:i}); return {rows:999,columns:0,bytes:0}; }
      if(input.fallback) globalThis.CompressionStream = undefined;
      const options = {...input.options,columns:input.columns,signal,onProgress};
      if(input.fontBytes) options.fonts = {regular:await api.PdfFont.create(input.fontBytes,{signal})};
      return {csv:api.writeCsvTo,xlsx:api.writeXlsxTo,pdf:api.writePdfTo}[input.format](rows,sink,options);
    });`;
  function worker() {
    const bootstrap = classic ? classicScript + '\n' + executionScript + '\nconst api=OfficeIMO,execution=OfficeIMO;\n' + workerBody
      : 'import * as api from ' + JSON.stringify(moduleBase + '/index.js') + ';\nimport * as execution from ' + JSON.stringify(moduleBase + '/execution/index.js') + ';\n' + workerBody;
    const url = URL.createObjectURL(new Blob([bootstrap], { type: 'text/javascript' }));
    const owned = new Worker(url, classic ? undefined : { type: 'module' }), listeners = new Set();
    const observed = { requests: 0, maxRows: 0, maxChunk: 0, progress: 0 };
    const endpoint = {
      postMessage(data) { if (data.kind === 'rows') observed.maxRows = Math.max(observed.maxRows,data.rows.length); owned.postMessage(data); },
      addEventListener(type, listener) {
        const wrapped = event => {
          if(type === 'message') {
            if(event.data.kind === 'requestRows') observed.requests++;
            if(event.data.kind === 'output') observed.maxChunk=Math.max(observed.maxChunk,event.data.bytes.length);
            if(event.data.kind === 'progress') observed.progress++;
          }
          listener(event);
        };
        listeners.add({type,listener,wrapped}); owned.addEventListener(type,wrapped);
      },
      removeEventListener(type,listener) { const entry=[...listeners].find(item=>item.type===type&&item.listener===listener); if(entry){owned.removeEventListener(type,entry.wrapped);listeners.delete(entry);} }
    };
    return { endpoint, observed, listeners, release() { owned.terminate(); URL.revokeObjectURL(url); } };
  }
  async function withWorker(run) { const owned=worker(); try { return await run(owned); } finally { owned.release(); } }
  async function save(name, chunks, contract = {}) {
    const data = new Uint8Array(await new Blob(chunks).arrayBuffer()); let text='';
    for(let first=0;first<data.length;first+=8192) text+=String.fromCharCode(...data.subarray(first,first+8192));
    await writeExecutionFixture((classic?'classic-':'esm-')+name,btoa(text),JSON.stringify(contract));
  }
  async function rejection(promise, expected, label) { let error; try { await promise; } catch(caught) { error=caught; } require(error&&expected(error),label+': '+error); checks.push(label); }
  function rows() { return [
    [new api.ExportCell('Łódź',{text:'City Łódź',presentation:{bold:true},link:{target:'https://example.com/city'}}),12.5,true],
    ['second',7,false]
  ]; }
  for(const format of ['csv','xlsx','pdf']) for(const fallback of format==='xlsx'?[false,true]:[false]) {
    await withWorker(async owned => {
      const chunks=[], sourceFont=fontBytes.slice();
      const completed=await execution.executeWorkerExport(owned.endpoint,rows(),{async write(bytes){chunks.push(bytes);await new Promise(resolve=>setTimeout(resolve,1));}},
        {input:{format,columns,fallback,options:format==='pdf'?{pageNumbers:false,compression:false}:{},...(format==='pdf'?{fontBytes:sourceFont}:{})},batchRows:1,maxBatchCells:3,maxChunkBytes:4096});
      require(completed.rows===2&&completed.columns===3,'Worker completion counts differ');
      require(chunks.reduce((size,chunk)=>size+chunk.length,0)===completed.bytes,'Worker completion bytes differ');
      require(owned.observed.maxRows<=1&&owned.observed.maxChunk<=4096&&owned.listeners.size===0,'Worker message/cleanup bounds differ');
      require(sourceFont.byteLength===fontBytes.byteLength&&sourceFont.every((value,index)=>value===fontBytes[index]),'Caller font bytes were detached or changed');
      await save('worker-'+format+(fallback?'-fallback':'')+'.'+format,chunks,{required:['Name','City Łódź','12.5','second','7']});
      checks.push(format+(fallback?'-fallback':'')+'-bytes-values-and-ownership');metrics.push({format,fallback,...owned.observed,bytes:completed.bytes});
    });
  }
  await withWorker(async owned => {
    let produced=0,release,ready;
    const gate=new Promise(resolve=>{release=resolve;}),started=new Promise(resolve=>{ready=resolve;});let writes=0;
    function* source(){for(let i=0;i<100;i++){produced++;yield ['x'.repeat(1000)+i];}}
    const pending=execution.executeWorkerExport(owned.endpoint,source(),{async write(){if(++writes===1){ready();await gate;}}},
      {input:{format:'csv',columns:[{header:'Value'}]},batchRows:3,maxBatchCells:3,maxChunkBytes:8192});
    await started;const stopped=produced;await new Promise(resolve=>setTimeout(resolve,10));
    require(stopped<100&&produced===stopped,'Destination backpressure failed to stop source production');release();
    require((await pending).rows===100&&owned.listeners.size===0,'Backpressure completion/cleanup failed');checks.push('acknowledged-output-stops-source');
  });
  for(const phase of ['source','sink']) await withWorker(async owned => {
    const controller=new AbortController(),original=Error(phase+' browser cancelled');let ready,returned=0;
    const started=new Promise(resolve=>{ready=resolve;});
    const source=phase==='source'?{[Symbol.asyncIterator]:()=>({next(){ready();return new Promise(()=>{});},return(){returned++;return{done:true};}})}:[['x'.repeat(40000)]];
    const pending=execution.executeWorkerExport(owned.endpoint,source,{write(){if(phase==='sink'){ready();return new Promise(()=>{});}}},
      {signal:controller.signal,input:{format:'csv',columns:[{header:'Value'}]}});
    await started;controller.abort(original);await rejection(pending,error=>error===original,'cancel-pending-'+phase);
    await new Promise(resolve=>setTimeout(resolve,0));
    require(owned.listeners.size===0&&(phase!=='source'||returned===1),'Cancellation retained listeners/producer');
    const reused=await execution.executeWorkerExport(owned.endpoint,[['reused']],{write(){}},{input:{format:'csv',columns:[{header:'Value'}]}});
    require(reused.rows===1,'Cancelled worker endpoint could not be reused');checks.push('reuse-after-'+phase+'-cancellation');
  });
  await withWorker(async owned => {
    await rejection(execution.executeWorkerExport(owned.endpoint,[],{write(){}},{input:{mode:'failure'}}),error=>error.code==='RESOURCE_LIMIT','remote-error-code');
    let last=0;
    const completed=await execution.executeWorkerExport(owned.endpoint,[],{write(){}},{input:{mode:'progress'},onProgress(progress){last=progress.rows;}});
    require(completed.rows===999&&last===999&&owned.observed.progress<=2,'Progress queue did not coalesce to bounded acknowledged updates');checks.push('coalesced-progress');
    const original=Error('progress callback failed');
    await rejection(execution.executeWorkerExport(owned.endpoint,[],{write(){}},{input:{mode:'progress'},onProgress(){throw original;}}),error=>error===original,'progress-failure');
    await rejection(execution.executeWorkerExport(owned.endpoint,[],{write(){}},{input:{mode:'dispose'}}),error=>error.name==='AbortError','worker-handler-disposal');
    require(owned.listeners.size===0,'Terminal worker case retained listeners');
  });
  return {checks:checks.length,labels:checks,metrics};
}

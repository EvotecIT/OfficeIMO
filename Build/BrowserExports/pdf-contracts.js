/* Product-specific assertions; browser lifetime and installation belong to HtmlTinkerX. */
globalThis.runPdfContracts = async function ({ regular, bold, symbols, japanese, workerScript, scale }) {
  const { writePdf, writePdfTo, PdfFont, ExportCell } = OfficeIMO;
  const bytes = base64 => Uint8Array.from(atob(base64), c => c.charCodeAt(0));
  const fonts = { regular: new PdfFont(bytes(regular)), bold: new PdfFont(bytes(bold)) };
  const columns = [{ header: 'Name', key: 'name', groups: ['Report'] }, { header: 'Amount', key: 'amount', alignment: 'right', groups: ['Report'] }, { header: 'State', key: 'state', groups: ['Report'] }];
  const rows = Array.from({length: 120}, (_, i) => ({ name: 'Row' + String(i).padStart(6, '0'), amount: i + .5,
    state: new ExportCell('Ready', {text: i === 0 ? 'Łódź — Zażółć gęślą jaźń' : 'Ready', presentation: { background: i === 0 ? 'C6EFCE' : 'ffffff' }}) }));
  const reports = [];
  async function save(name, blob, required, extra = {}) {
    const data = new Uint8Array(await blob.arrayBuffer()); let encoded = '';
    for (let offset = 0; offset < data.length; offset += 32768) encoded += String.fromCharCode(...data.subarray(offset, offset + 32768));
    await writePdfFixture(name + '.pdf', btoa(encoded), JSON.stringify({required, ...extra}));
  }
  await save('empty',await writePdf([],{columns:columns.map(({groups,...column})=>column),fonts,includeHeader:false,pageNumbers:false}),[]);
  const modified = bytes(regular), metricView = new DataView(modified.buffer);
  for (let i=0;i<metricView.getUint16(4);i++) {
    const p=12+i*16;
    if(String.fromCharCode(...modified.subarray(p,p+4))==='hhea') {
      const at=metricView.getUint32(p+8);metricView.setInt16(at+4,4096);metricView.setInt16(at+6,-1024);
    }
  }
  await save('mixed-metrics',await writePdf([[new ExportCell('Tall',{presentation:{bold:true}})],['Next']],{
    columns:[{header:'Heading'}],fonts:{regular:fonts.regular,bold:new PdfFont(modified)},title:'Title'}),['Title','Heading','Tall','Next']);
  await save('page-number-size',await writePdf([['Value']],{columns:[{header:'Heading'}],fontSize:36,pageSize:{width:700,height:900},
    margins:{left:36,right:36,top:36,bottom:80},pageFooter:'End'}),['Heading','Value','Page 1 of','End']);
  await save('afterword',await writePdf([['Data']],{columns:[{header:'Table heading'}],pageSize:'A5',pageHeader:'Report',pageFooter:'Footer',
    messageBottom:'Afterword '.repeat(1000)}),['Data','Afterword'],{firstPageOnly:['Table heading'],repeated:['Report','Footer']});
  await save('symbols',await writePdf([['🂡♟']],{columns:[{header:''}],fonts:{regular:new PdfFont(bytes(symbols))},pageNumbers:false}),['🂡♟']);
  await save('japanese',await writePdf([['東京 大阪 日本語']],{columns:[{header:''}],fonts:{regular:new PdfFont(bytes(japanese))},pageNumbers:false}),['東京 大阪 日本語']);
  for (const compression of [true, false]) {
    const start = scale ? performance.now() : undefined;
    await save('report-' + compression, await writePdf(rows, { columns, fonts, compression, title: 'Raport Łódź',
      footer: {values: ['Totals'], totals: {amount:'sum'}}, alternateRowColor: 'F3F4F6', pageHeader: 'OfficeIMO', pageFooter: 'Confidential' }),
      ['Raport Łódź', 'Łódź — Zażółć gęślą jaźń', 'Totals', '7200'], {rows:120, repeated:['Report','Name','Amount','State']});
    if (start !== undefined) reports.push({name:'report-' + compression, milliseconds:performance.now()-start});
  }
  await save('empty-metadata', await writePdf([['Row000000', 1]], { columns: [{ header: 'Name' }, { header: 'Amount' }],
    title: '', messageTop: '', messageBottom: '', pageHeader: '', pageFooter: () => '', pageNumbers: false, margins: 0,
    limits: { maxCells: 4 } }), ['Name', 'Amount', 'Row000000', '1'],
    { rows: 1, repeated: ['Name', 'Amount'] });
  const compressor = globalThis.CompressionStream;
  try {
    globalThis.CompressionStream = undefined;
    await save('fallback', await writePdf([['Fallback €']], {columns:[{header:'Heading'}]}), ['Fallback €','Heading']);
  } finally { globalThis.CompressionStream = compressor; }
  await save('spans', await writePdf(rows, {columns, fonts,
    headerRows: [[{value:'Name',rowSpan:2},{value:'Metrics',columnSpan:2},null], [null,{value:'Amount'},{value:'State'}]],
    footer: { rows: [[{value:'Totals',rowSpan:2},{value:'7200'},{value:'Approved'}], [null,{value:'Finance',columnSpan:2},null]] },
    pageSize:'LETTER',orientation:'landscape' }), ['Totals','7200','Approved','Finance'], {rows:120,repeated:['Name','Metrics','Amount','State']});
  const long = 'ŁódźABCDEFGHIJKLMNOPQRSTUVWXYZ'.repeat(500);
  await save('long', await writePdf([[long]], {fonts,columns:[{header:'Long text'}],columnWidths:[200],pageSize:'A5',pageNumbers:false}), [], {repeated:['Long text'],bodyText:long});
  let writes = 0, sourceRows = 0, firstByteRows, returned = false;
  const chunks = [];
  async function* paged() {
    try { for (let first=0;first<rows.length;first+=16) { await new Promise(resolve=>setTimeout(resolve,1)); for (const row of rows.slice(first,first+16)) {sourceRows++;yield row;} } }
    finally {returned=true;}
  }
  await writePdfTo(paged(), {async write(chunk) { if (!writes) firstByteRows=sourceRows; writes++;chunks.push(chunk.slice());await new Promise(resolve=>setTimeout(resolve,1)); }}, {columns,fonts});
  if (!returned || firstByteRows >= rows.length) throw Error('PDF must stream pages before source exhaustion.');
  await save('slow-paged',new Blob(chunks),['Łódź — Zażółć gęślą jaźń'],{rows:120,repeated:['Name','Amount','State']});
  reports.push({name:'slow-paged',writes,firstByteRows});
  const controller = new AbortController(); let stopped=false;
  async function* cancellable() {try {for (let i=0;i<10000;i++) yield rows[i%rows.length];}finally {stopped=true;}}
  try {await writePdfTo(cancellable(),{write(){controller.abort('pdf cancelled');}},{columns,fonts,signal:controller.signal});throw Error('Cancellation ignored');}
  catch(error) {if(error !== 'pdf cancelled') throw error;}
  if(!stopped) throw Error('Cancellation did not return source.');
  for(const limits of [{maxRows:1},{maxPages:0},{maxCellCharacters:2},{maxOutputBytes:100}]) {
    try {await writePdf(rows,{columns,fonts,limits});throw Error('Resource limit ignored');}
    catch(error) {if(!String(error).includes('exceeded')) throw error;}
  }
  for(const scenario of ['native','fallback','cancel']) {
    const bootstrap = `let controller,ack;
    onmessage=async({data})=>{
      if(data.ack){const done=ack;ack=undefined;done?.();return;}
      if(data.cancel){controller.abort('worker cancelled');return;}
      controller=new AbortController();try{
      if(data.fallback)globalThis.CompressionStream=undefined;
      const fonts={regular:new OfficeIMO.PdfFont(Uint8Array.from(atob(data.regular),c=>c.charCodeAt(0)))};
      async function* rows(){for(let first=0;first<300;first+=32){await new Promise(r=>setTimeout(r,0));for(let i=first;i<Math.min(300,first+32);i++)yield ['Row'+String(i).padStart(6,'0'),'Łódź'];}}
      await OfficeIMO.writePdfTo(rows(),{write(bytes){const chunk=bytes.slice();return new Promise(resolve=>{ack=resolve;postMessage({chunk},[chunk.buffer]);});}},
        {columns:[{header:'Name'},{header:'City'}],fonts,signal:controller.signal});
      postMessage({done:true});
    }catch(error){postMessage({error:String(error)})}};`;
    const url=URL.createObjectURL(new Blob([workerScript,'\n',bootstrap],{type:'text/javascript'})),worker=new Worker(url);
    let timeout;
    try {
      const chunks=[];
      const result=await new Promise((resolve,reject)=>{
        timeout=setTimeout(()=>reject(Error('PDF worker timed out')),30000);
        worker.onerror=e=>reject(Error(e.message));worker.onmessage=async({data})=>{
          if(data.chunk){chunks.push(data.chunk);if(scenario==='cancel')worker.postMessage({cancel:true});await new Promise(r=>setTimeout(r,1));worker.postMessage({ack:true});}
          else if(data.error){if(scenario==='cancel'&&data.error==='worker cancelled')resolve({cancelled:true});else reject(Error(data.error));}
          else if(data.done)resolve({blob:new Blob(chunks)});
        };
        worker.postMessage({regular,fallback:scenario==='fallback'});
      });
      if(scenario==='cancel'){if(!result.cancelled)throw Error('Worker cancellation ignored');}
      else await save('worker-'+scenario,result.blob,['Łódź'],{rows:300,repeated:['Name','City']});
    } finally {clearTimeout(timeout);worker.terminate();URL.revokeObjectURL(url);}
  }
  if(globalThis.DataTable) {
    const metadata={title:'Default title',messageTop:'Default above',messageBottom:'Default below'};
    document.body.innerHTML='<table id="pdf-table"><thead><tr><th rowspan="2">Name</th><th colspan="2">Metrics</th></tr><tr><th>Amount</th><th>State</th></tr></thead><tfoot><tr><th rowspan="2">Totals</th><th>7200</th><th>Approved</th></tr><tr><th colspan="2">Finance</th></tr></tfoot></table>';
    const table=new DataTable('#pdf-table',{data:rows.map(r=>[r.name,r.amount,'Łódź']),paging:true,columns:[{title:'Name'},{title:'Amount'},{title:'State'}]});
    try {
      await save('datatables',await OfficeIMO.exportDataTable(DataTable,table,'pdf',{pdf:{fonts,title:'Table PDF'},exportOptions:{modifier:{order:'index',search:'none'}}}),
        ['Table PDF','Łódź','Metrics','Totals','7200','Approved','Finance'],{rows:120,repeated:['Name','Amount','State']});
      let delivered;
      OfficeIMO.registerDataTablesButtons(DataTable,{pdf:{fonts,...metadata},save:async blob=>{delivered=blob;},onError:error=>{throw error;}});
      await new Promise(resolve=>DataTable.ext.buttons.officeimoPdf.action(null,table,null,{title:'Button report',pageSize:'LETTER',orientation:'landscape',exportOptions:{modifier:{order:'index',search:'none'}}},resolve));
      if(!delivered) throw Error('PDF button failed to deliver its file.');
      await save('datatables-button',delivered,['Button report','Łódź','Metrics','Totals','Approved','Finance'],{rows:120});
      for (const location of ['registration','button']) {
        const host={Buttons:DataTable.Buttons,ext:{buttons:{}}}, replacement={fonts,pageNumbers:false,
          headerRows:[[{value:'Replacement heading'},{value:'Amount'},{value:'State'}]],
          footer:{rows:[[{value:'Replacement footer'},{value:'7200'},{value:'Approved'}]]}};
        OfficeIMO.registerDataTablesButtons(host,{...(location==='registration'?{pdf:replacement}:{}),
          save:blob=>{delivered=blob;},onError:error=>{throw error;}});
        for (const suppressed of ['header','footer','both']) {
          delivered=undefined;
          await new Promise(resolve=>host.ext.buttons.officeimoPdf.action(null,table,null,{
            ...(suppressed!=='footer'?{header:false}:{}),...(suppressed!=='header'?{footer:false}:{}),exportOptions:{modifier:{order:'index',search:'none'}},
            ...(location==='button'?{officeimo:{pdf:replacement}}:{})},resolve));
          if(!delivered)throw Error('Replacement-heading/footer suppression did not deliver.');
          const required=['Łódź'],forbidden=['Metrics','Totals','Finance'];
          (suppressed==='footer'?required:forbidden).push('Replacement heading');
          (suppressed==='header'?required:forbidden).push('Replacement footer');
          await save('datatables-suppression-'+location+'-'+suppressed,delivered,required,{forbidden,rows:120});
        }
        if(replacement.headerRows[0][0].value!=='Replacement heading'||replacement.footer.rows[0][0].value!=='Replacement footer')
          throw Error('Native suppression mutated registration/button options.');
      }
      for(const key of ['omitted','title','messageTop','messageBottom','all']) {
        delivered=undefined;
        const overrides=key==='omitted'?{}:key==='all'?{title:null,messageTop:null,messageBottom:null}:{[key]:null};
        await new Promise(resolve=>DataTable.ext.buttons.officeimoPdf.action(null,table,null,
          {...overrides,exportOptions:{modifier:{order:'index',search:'none'}}},resolve));
        if(!delivered)throw Error('PDF metadata button failed to deliver its file.');
        const required=['Łódź'],forbidden=[];
        for(const name of Object.keys(metadata))(key===name||key==='all'?forbidden:required).push(metadata[name]);
        await save('datatables-metadata-'+key,delivered,required,{forbidden,rows:120});
      }
      for(const key of ['title','messageTop','messageBottom','empty']) {
        delivered=undefined;
        const overrides=key==='empty'?{title:'',messageTop:'',messageBottom:''}:{[key]:key==='title'?()=>null:'*'};
        await new Promise(resolve=>DataTable.ext.buttons.officeimoPdf.action(null,table,null,
          {...overrides,exportOptions:{modifier:{order:'index',search:'none'}}},resolve));
        if(!delivered)throw Error('Resolved-empty PDF metadata button failed to deliver its file.');
        const required=['Łódź'],forbidden=[];
        for(const name of Object.keys(metadata))(key===name||key==='empty'?forbidden:required).push(metadata[name]);
        await save('datatables-metadata-resolved-'+key,delivered,required,{forbidden,rows:120});
      }
    } finally {table.destroy(true);}
    const deepRows=rows.slice(0,3).map(r=>[r.name,r.amount,'Łódź']);
    const exportOptions={modifier:{order:'index',search:'none'}};
    const sourceFooter='<tfoot><tr><th rowspan="2">Totals</th><th>4.5</th><th>Approved</th></tr><tr><th colspan="2">Finance</th></tr></tfoot>';
    document.body.innerHTML='<table id="deep-header"><thead>'+Array.from({length:17},()=>'<tr><th>Unused heading</th><th>Unused heading</th><th>Unused heading</th></tr>').join('')+'</thead>'+sourceFooter+'</table>';
    const deepHeader=new DataTable('#deep-header',{data:deepRows,order:[],columns:[{title:'Name'},{title:'Amount'},{title:'State'}]});
    try {
      await save('datatables-no-header',await OfficeIMO.exportDataTable(DataTable,deepHeader,'pdf',{exportOptions,pdf:{fonts,includeHeader:false}}),
        ['Łódź','Totals','Approved','Finance'],{rows:3,forbidden:['Unused heading']});
      await save('datatables-replacement-header',await OfficeIMO.exportDataTable(DataTable,deepHeader,'pdf',{exportOptions,
        pdf:{fonts,headerRows:[[{value:'Replacement heading'},{value:'Amount'},{value:'State'}]]}}),
        ['Replacement heading','Łódź','Totals','Finance'],{rows:3,forbidden:['Unused heading']});
      let delivered,failure;
      await new Promise(resolve=>DataTable.ext.buttons.officeimoPdf.action(null,deepHeader,null,
        {header:false,exportOptions,officeimo:{save:blob=>{delivered=blob;},onError:error=>{failure=error;}}},resolve));
      if(failure||!delivered)throw failure??Error('Suppressed-header PDF button did not deliver.');
      await save('datatables-no-header-button',delivered,['Łódź','Totals','Finance'],{rows:3,forbidden:['Unused heading']});
    } finally {deepHeader.destroy(true);}
    document.body.innerHTML='<table id="deep-footer"><thead><tr><th>Name</th><th>Amount</th><th>State</th></tr></thead><tfoot>'+Array.from({length:17},()=>'<tr><th>Unused footer</th><th>Unused footer</th><th>Unused footer</th></tr>').join('')+'</tfoot></table>';
    const deepFooter=new DataTable('#deep-footer',{data:deepRows,order:[],columns:[{title:'Name'},{title:'Amount'},{title:'State'}]});
    try {
      await save('datatables-replacement-footer',await OfficeIMO.exportDataTable(DataTable,deepFooter,'pdf',{exportOptions,
        pdf:{fonts,footer:{values:['Replacement footer',42,'Approved']}}}),
        ['Name','Łódź','Replacement footer','42','Approved'],{rows:3,forbidden:['Unused footer']});
    } finally {deepFooter.destroy(true);}
    document.body.innerHTML='<table id="empty-metadata"><thead><tr><th>Name</th><th>Amount</th></tr></thead></table>';
    const emptyMetadata=new DataTable('#empty-metadata',{data:[['Row000000',1],['Row000001',2]],order:[],columns:[{title:'Name'},{title:'Amount'}]});
    try {
      let delivered,failure;
      await new Promise(resolve=>DataTable.ext.buttons.officeimoPdf.action(null,emptyMetadata,null,
        {title:'',messageTop:'',messageBottom:'',header:false,footer:false,exportOptions,
          officeimo:{limits:{maxCells:4},pdf:{fonts,pageNumbers:false},save:blob=>{delivered=blob;},onError:error=>{failure=error;}}},resolve));
      if(failure||!delivered)throw failure??Error('Empty-metadata PDF button did not deliver within the data-cell budget.');
      await save('datatables-empty-metadata-budget',delivered,['Row000000','Row000001'],{rows:2,forbidden:Object.values(metadata)});
    } finally {emptyMetadata.destroy(true);}
  }
  if(scale) for(const count of [10000,100000]) for(const width of [4,20]) for(const styled of [false,true]) {
    async function* source(){for(let i=0;i<count;i++)yield Array.from({length:width},(_,c)=>c===0?'Row'+String(i).padStart(6,'0'):styled?new ExportCell(i+c,{text:String(i+c),presentation:{background:i%2?'F3F4F6':'ffffff'}}):i+c);}
    const start=performance.now(),blob=await writePdf(source(),{columns:Array.from({length:width},(_,c)=>({header:'Column'+c})),pageSize:'A3',orientation:'landscape',fontSize:7,pageNumbers:false,limits:{maxPages:10000}});
    reports.push({name:'scale',rows:count,columns:width,styled,milliseconds:performance.now()-start,bytes:blob.size});
    await save('scale-'+count+'-'+width+'-'+styled,blob,[],{rows:count,columns:width,repeated:['Column0']});
  }
  return {reports,cancellation:true,limits:true,workers:2,workerCancellation:true,compressionFallback:true};
};

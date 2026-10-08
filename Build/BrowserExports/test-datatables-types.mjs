// Opt-in consumer proof. Third-party declarations stay in caller-owned validation folders.
import { spawnSync } from 'node:child_process';
import { mkdir, writeFile, readFile, copyFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('../../OfficeIMO.JavaScript/',import.meta.url));
const output = process.argv[2] && resolve(process.argv[2]);
if (!output || !process.env.npm_execpath) throw new Error('Use npm run test:datatables-types -- <packed-evidence-directory>.');
const archive = JSON.parse(await readFile(join(output,'consumer-report.json'),'utf8')).archive;
const consumer = `import DataTable from 'datatables.net';
import 'datatables.net-buttons';
import {createDataTablesExport,exportDataTable,registerDataTablesButtons} from '@evotecit/officeimo/integrations/datatables';
const table = new DataTable('#grid');
createDataTablesExport(DataTable,table,{exportOptions:{columns:':visible'}});
exportDataTable(DataTable,table,'xlsx');
exportDataTable(DataTable,table,'pdf',{pdf:{pageSize:'A4',orientation:'landscape'}});
registerDataTablesButtons(DataTable);
`;
const reports = [];
for (const [name,dt,buttons] of [['bundled','2.3.7','3.2.6'],['current','3.1.3','4.1.2']]) {
  const folder = join(output,'types-'+name); await mkdir(folder,{recursive:true});
  await writeFile(join(folder,'package.json'),'{"private":true,"type":"module"}\n');
  const installed = spawnSync(process.execPath,[process.env.npm_execpath,'install','--ignore-scripts','--no-audit','--no-fund','--save-exact',join(output,archive),
    'datatables.net@'+dt,'datatables.net-buttons@'+buttons,'@types/jquery@3.5.34'],{cwd:folder,encoding:'utf8',timeout:120000});
  if (installed.status !== 0) throw new Error(installed.error?.message ?? installed.stdout+installed.stderr);
  await writeFile(join(folder,'consumer.mts'),consumer);
  const flags = [join(root,'node_modules/typescript/bin/tsc'),'--strict','--exactOptionalPropertyTypes','--noUncheckedIndexedAccess','--noEmit',
    '--target','ES2022','--module','ESNext','--moduleResolution','Bundler','--lib','ES2022,DOM','consumer.mts'];
  const full = spawnSync(process.execPath,flags,{cwd:folder,encoding:'utf8',timeout:30000});
  await writeFile(join(folder,'declaration-check.log'),full.stdout+full.stderr);
  // DT3/Buttons4 currently duplicate upstream index signatures. Never conceal a consumer error:
  // checking the consumer with skipLibCheck still validates every adapter call and assignability.
  if (full.status !== 0 && name === 'bundled') throw new Error(full.stdout+full.stderr);
  const checked = spawnSync(process.execPath,[...flags,'--skipLibCheck'],{cwd:folder,encoding:'utf8',timeout:30000});
  if (checked.status !== 0) throw new Error(checked.stdout+checked.stderr);
  await copyFile(join(folder,'package-lock.json'),join(output,'types-'+name+'-lock.json'));
  reports.push({name,dataTables:dt,buttons,consumerPassed:true,fullDeclarationsPassed:full.status===0,compiler:'5.9.3',resolution:'Bundler'});
}
await writeFile(join(output,'datatables-types.json'),JSON.stringify(reports,null,2)+'\n');
console.log(JSON.stringify(reports));

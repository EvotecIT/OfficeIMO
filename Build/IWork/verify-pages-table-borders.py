"""Independent ZIP/XML readback of the controlled Pages border corpus."""
import hashlib,json,sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile
decode,receipt,*inputs=map(Path,sys.argv[1:])
expected=json.loads(decode.read_text())
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
w='{'+ns['w']+'}'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
records=[]
for path in inputs:
 with ZipFile(path) as z:
  assert z.testzip() is None
  document=ET.fromstring(z.read('word/document.xml'))
 tables=document.findall('.//w:tbl',ns);assert len(tables)==1
 rows=tables[0].findall('w:tr',ns);assert len(rows)==4
 count=bold_cells=0
 for cell in expected['cells']:
  r,c=cell['row'],cell['column']
  cells=rows[r-1].findall('w:tc',ns);assert len(cells)==4
  target=cells[c-1]
  text=''.join(t.text or '' for t in target.findall('.//w:t',ns))
  assert text==cell['value'],(path.name,cell['address'],text,cell['value'])
  for side,value in cell['edges'].items():
   edge=target.find('w:tcPr/w:tcBorders/w:'+side,ns);assert edge is not None
   width=value['widthPoints'];style=edge.get(w+'val')
   if width==0:assert style=='nil',(path.name,cell['address'],side,style)
   else:
    assert style=='single' and int(edge.get(w+'sz'))==width*8
    assert edge.get(w+'color').upper()==value['rgb']
   count+=1
  text_runs=[run for run in target.findall('.//w:r',ns) if any(t.text for t in run.findall('w:t',ns))]
  wanted_bold=r in (1,4) or c==1
  for run in text_runs:
   bold=run.find('w:rPr/w:b',ns)
   actual=bold is not None and bold.get(w+'val','true') not in ('0','false','off')
   assert actual==wanted_bold,(path.name,cell['address'],'bold',actual,wanted_bold)
  bold_cells+=int(bool(text_runs) and wanted_bold)
 assert count==64 and bold_cells==10
 records.append({'file':path.name,'bytes':path.stat().st_size,'sha256':sha(path),
  'values':16,'edges':count,'boldCells':bold_cells,'emptyD3':True})
receipt.write_text(json.dumps({'method':'Python stdlib ZIP/XML; no OfficeIMO loads',
 'decodeSha256':sha(decode),'artifacts':records},indent=2)+'\n')
print('PASS:',len(records),'DOCX files; all 16 values, 64 edges, 10 bold cells and empty D3')

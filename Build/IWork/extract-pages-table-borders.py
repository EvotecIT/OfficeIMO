"""Verify a controlled native Pages table against Apple's DOCX export.

Opt-in evidence: numbers-parser 4.19.0 and a pinned MIT KN/TST schema.
Checks the imported-and-natively-edited fixture only, with no OfficeIMO loads.
Every grid segment is explicitly selected; inherited border rendering and
arbitrary cell storage envelopes are outside this controlled reference.
"""
import argparse,hashlib,json,struct
from importlib.metadata import version
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from google.protobuf import descriptor_pb2,descriptor_pool,message_factory
from numbers_parser.model import get_storage_buffers_for_row
from iwork_protobuf_evidence import archive_messages
parser=argparse.ArgumentParser(description=__doc__)
parser.add_argument('source',type=Path)
parser.add_argument('descriptor',type=Path)
parser.add_argument('native_docx',type=Path)
parser.add_argument('output',type=Path)
args=parser.parse_args()
source,descriptor,docx=args.source,args.descriptor,args.native_docx
assert version('numbers-parser')=='4.19.0'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
assert sha(source)=='d6fefca93e07ab4c2f8557f91ad6dd34bca5864c0372733751d74db992fcc65d'
assert sha(descriptor)=='2408cf25920c39fe59d500a0b580bd64bf10a24630c81ca2d543e64322853917'
pool=descriptor_pool.DescriptorPool()
for f in descriptor_pb2.FileDescriptorSet.FromString(descriptor.read_bytes()).file:pool.Add(f)
graph={}
with ZipFile(source) as z:
 assert z.testzip() is None
 for _,h,i,b in archive_messages(z):graph.setdefault((h.identifier,i.type),[]).append(b)
def decode(identifier,kind,name):
 data=graph[(identifier,kind)];assert len(data)==1
 m=message_factory.GetMessageClass(pool.FindMessageTypeByName(name)).FromString(data[0])
 assert not m.FindInitializationErrors();return m
assert [i for i,k in graph if k==6001]==[1732826]
m=decode(1732826,6001,'TST.TableModelArchive')
assert (m.number_of_rows,m.number_of_columns)==(4,4)
assert (m.number_of_header_rows,m.number_of_header_columns,m.number_of_footer_rows)==(1,1,1)
style=decode(m.table_style.identifier,6003,'TST.TableStyleArchive')
assert style.table_properties.HasField('banded_rows') and not style.table_properties.banded_rows
assert style.table_properties.HasField('auto_resize') and not style.table_properties.auto_resize
parent=decode(style.super.parent.identifier,6003,'TST.TableStyleArchive')
inactive_defaults={}
for name in ['table_body_horizontal_border_stroke','table_body_vertical_border_stroke',
             'table_body_horizontal_stroke','table_body_vertical_stroke']:
 stroke=getattr(parent.table_properties,name)
 assert parent.table_properties.HasField(name) and stroke.width==1 and stroke.join==1 and stroke.pattern.type==1
 inactive_defaults[name]={'widthPoints':stroke.width,'join':'RoundJoin','pattern':'Solid'}
sidecar=decode(m.stroke_sidecar.identifier,6305,'TST.StrokeSidecarArchive')
assert (sidecar.max_order,sidecar.row_count,sidecar.column_count)==(2,4,4)
edges,layers={},[]
for name in ['left_column_stroke_layers','right_column_stroke_layers','top_row_stroke_layers','bottom_row_stroke_layers']:
 refs=getattr(sidecar,name);assert len(refs)==4
 horizontal='row' in name;increment=int(name.startswith(('right','bottom')))
 for ref in refs:
  layer=decode(ref.identifier,6306,'TST.StrokeLayerArchive');line=layer.row_column_index+increment;runs=[]
  for run in layer.stroke_runs:
   s=run.stroke;assert s.pattern.type in (1,2) and s.width in (1,2) and s.color.a==1
   assert 0<=run.origin<4 and 0<run.length<=4-run.origin and run.order<=sidecar.max_order
   visible=s.pattern.type==1
   rgb=''.join(f'{round(v*255):02X}' for v in [s.color.r,s.color.g,s.color.b]) if visible else '000000'
   value={'widthPoints':s.width if visible else 0,'rgb':rgb}
   runs.append({'origin':run.origin,'length':run.length,'order':run.order,**value})
   for pos in range(run.origin,run.origin+run.length):
    key=(horizontal,line,pos);prev=edges.get(key)
    if prev is None or prev[0]<run.order:edges[key]=(run.order,value)
    elif prev[0]==run.order:assert prev[1]==value
  layers.append({'id':ref.identifier,'side':name,'line':line,'runs':runs})
assert len(edges)==40 # Every rectangular grid segment has an explicit source edit.
cat=decode(m.base_data_store.rich_text_table.identifier,6005,'TST.TableDataList')
texts={}
for entry in cat.entries:
 rich=decode(entry.rich_text_payload.identifier,6218,'TST.RichTextPayloadArchive')
 texts[entry.key]=''.join(decode(rich.storage.identifier,2001,'TSWP.StorageArchive').text)
tile_ref=m.base_data_store.tiles.tiles[0].tile.identifier
tile=decode(tile_ref,6002,'TST.Tile');assert tile.last_saved_in_BNC and tile.storage_version==5
values={}
for row in tile.rowInfos:
 buffers=get_storage_buffers_for_row(row.cell_storage_buffer,row.cell_offsets,4,row.has_wide_offsets)
 assert len(buffers)==4
 for col,b in enumerate(buffers):
  assert b[0]==5 and b[1] in (0,9)
  flags=struct.unpack_from('<I',b,8)[0]
  if b[1]==9:
   assert flags==0x21070 and len(b)==32 # Controlled rich-text-only cell envelope.
   value=texts[struct.unpack_from('<I',b,12)[0]]
  else:
   assert flags==0x60 and len(b)==20;value=''
  values[(row.tile_row_index,col)]=value
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
w='{'+ns['w']+'}'
with ZipFile(docx) as z:
 assert z.testzip() is None
 document=ET.fromstring(z.read('word/document.xml'))
tables=document.findall('.//w:tbl',ns);assert len(tables)==1
rows=tables[0].findall('w:tr',ns);assert len(rows)==4
sides={'left':(False,0,0),'right':(False,0,1),'top':(True,0,0),'bottom':(True,1,0)}
cells=[]
for r,row in enumerate(rows):
 native=row.findall('w:tc',ns);assert len(native)==4
 for c,cell in enumerate(native):
  value=values[(r,c)];assert ''.join(t.text or '' for t in cell.findall('.//w:t',ns))==value
  expected={}
  for side,(horizontal,rd,cd) in sides.items():
   line=r+rd if horizontal else c+cd;pos=c if horizontal else r
   edge=edges[(horizontal,line,pos)][1];expected[side]=edge
   declaration=cell.find('w:tcPr/w:tcBorders/w:'+side,ns);assert declaration is not None
   border_style=declaration.attrib[w+'val']
   if edge['widthPoints']==0:assert border_style=='nil'
   else:
    assert border_style=='single' and int(declaration.attrib[w+'sz'])==edge['widthPoints']*8
    assert declaration.attrib[w+'color'].upper()==edge['rgb']
  cells.append({'row':r+1,'column':c+1,'address':chr(65+c)+str(r+1),'value':value,'edges':expected})
assert len(cells)==16 and sum(len(c['edges']) for c in cells)==64
out={'sourceSha256':sha(source),'nativeDocxSha256':sha(docx),'descriptorSha256':sha(descriptor),'model':1732826,
 'table':m.table_name,'rows':4,'columns':4,'headerRows':1,'headerColumns':1,'footerRows':1,
 'bandedRows':False,'autoResizeRows':False,'styleChain':[m.table_style.identifier,style.super.parent.identifier],
 'sidecar':m.stroke_sidecar.identifier,'inactiveRoundedDefaults':inactive_defaults,'layers':layers,'cells':cells}
args.output.write_text(json.dumps(out,indent=2)+'\n')
print('PASS: 16 native source values, 64 edges and explicit disabled banding/row auto-resize match Pages DOCX.')

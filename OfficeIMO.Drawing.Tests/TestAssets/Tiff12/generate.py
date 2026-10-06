"""Run the isolated native producer and preserve hashes of independently decoded samples."""
import hashlib, json, pathlib, subprocess, sys
root=pathlib.Path(__file__).resolve().parent
cases=[]
for kind in range(7):
 for compression in [1,5,8,32773,7]:
  if kind==5 and compression!=7 or kind in (3,4) and compression==7: continue
  for endian in [0,1]:
   for planar in [1,2]:
    if kind==5 and planar==2: continue
    for tiled in [0,1]:
     name=f'k{kind}-c{compression}-be{endian}-p{planar}-t{tiled}.tif'
     if len(sys.argv)>2 and kind<int(sys.argv[2]):
      cases.append(dict(file=name,kind=kind,compression=compression,width=35,height=19));continue
     subprocess.run([sys.argv[1],str(root/name),str(kind),str(compression),str(endian),str(planar),str(tiled),str(tiled)],check=True)
     cases.append(dict(file=name,kind=kind,compression=compression,width=35,height=19))
(root/'manifest.json').write_text(json.dumps(cases,indent=2)+'\n')
(root/'sha256.json').write_text(json.dumps({p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(root.iterdir()) if p.suffix in ('.tif','.raw')},indent=2)+'\n')
print(len(cases),'independent TIFF files')

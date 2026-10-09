"""Test-only independent producer: dbf 0.99.11; oracle: dbfread 2.0.7.
Fictional fixture contents use the repository MIT license.
"""
from pathlib import Path
from hashlib import sha256
import json
import dbf
from dbfread import DBF
root = Path(__file__).resolve().parent
table = dbf.Table(str(root / 'plain.dbf'), 'NAME C(40); AMOUNT N(12,2); ACTIVE L', dbf_type='db3', codepage='cp1252')
table.open(dbf.READ_WRITE)
table.append(('Café | <script> & **bold**', 1234.5, True))
table.append(('deleted', -2.75, False))
table.append(('', None, None))
dbf.delete(table[1])
table.close()
manifest = {'producer': 'dbf 0.99.11', 'oracle': 'dbfread 2.0.7',
    'records': list(DBF(str(root / 'plain.dbf'))),
    'sha256': sha256((root / 'plain.dbf').read_bytes()).hexdigest()}
(root / 'plain-manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')

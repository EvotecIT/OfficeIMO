"""Regenerate the sanitized fixture and native references using Pillow/LibTIFF."""
from pathlib import Path
import struct
from PIL import Image, features
root = Path(__file__).resolve().parent
source = root / 'ojpeg_single_strip_no_rowsperstrip.tiff'
data = bytearray(source.read_bytes())
e = '<' if data[:2] == b'II' else '>'
directory = struct.unpack_from(e + 'I', data, 4)[0]
count = struct.unpack_from(e + 'H', data, directory)[0]
records = [bytes(data[directory+2+i*12:directory+2+(i+1)*12]) for i in range(count)]
valid = [record for record in records if struct.unpack_from(e + 'HHI', record) != (0, 0, 0)]
assert len(records) - len(valid) == 2
next_ifd = bytes(data[directory+2+count*12:directory+6+count*12])
replacement = struct.pack(e + 'H', len(valid)) + b''.join(valid) + next_ifd
# Leave the former directory tail in place so every existing payload offset holds.
data[directory:directory+len(replacement)] = replacement
(root / 'ojpeg_single_strip_no_rowsperstrip_sanitized.tiff').write_bytes(data)
for path in sorted(root.glob('*.tiff')):
    Image.open(path).convert('RGBA').save(path.with_suffix('.png'))
print('Pillow', Image.__version__, 'LibTIFF', features.version('libtiff'))

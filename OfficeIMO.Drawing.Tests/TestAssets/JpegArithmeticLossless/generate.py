"""Generate arithmetic JPEG fixtures with external test-only native executables.
Usage: python3 generate.py /path/to/jpeg /path/to/djpeg /task/scratch [output]
The jpeg executable must use the revision/settings in prepare_oracle.py.
"""
import csv
import hashlib
import os
from pathlib import Path
import subprocess
import sys

oracle, turbo = map(lambda p: str(Path(p).resolve()), sys.argv[1:3])
scratch = Path(sys.argv[3]).resolve()
output = Path(sys.argv[4]).resolve() if len(sys.argv) > 4 else Path(__file__).resolve().parent
scratch.mkdir(parents=True, exist_ok=True)
output.mkdir(parents=True, exist_ok=True)


def run(command, env=None):
 result = subprocess.run(command, env=env, capture_output=True)
 # The native test driver can return zero on errors; stderr is authoritative too.
 assert result.returncode == 0 and not result.stderr, (command, result.stderr)


def pnm(path):
 data = path.read_bytes()
 at, tokens = 0, []
 while len(tokens) < 4:
  while data[at:at+1].isspace(): at += 1
  if data[at:at+1] == b'#':
   at = data.index(b'\n', at) + 1
   continue
  end = at
  while not data[end:end+1].isspace(): end += 1
  tokens.append(data[at:end]); at = end
 return tokens, data[at+1:]


def source(bits, components):
 maximum = (1 << bits) - 1
 values = [((x*193+y*791+c*3191) ^ (x*y*53)) & maximum
           for y in range(11) for x in range(19) for c in range(components)]
 data = bytes(values) if bits <= 8 else b''.join(v.to_bytes(2, 'big') for v in values)
 path = scratch / f'b{bits}-c{components}.pnm'
 path.write_bytes(f'P{5 if components == 1 else 6}\n19 11\n{maximum}\n'.encode()+data)
 return path, values


def encode(bits, components, predictor, point, restart, name, sampling=None):
 src, values = source(bits, components)
 env = dict(os.environ, OFFICEIMO_TEST_POINT=str(point), OFFICEIMO_TEST_PREDICTOR=str(predictor))
 jpeg = output / (name+'.jpg')
 options = ['-p', '-c', '-z', str(restart)] + (['-s', sampling] if sampling else [])
 run([oracle, '-a'] + options + [str(src), str(jpeg)], env)
 raw = jpeg.read_bytes(); sof = raw.index(b'\xff\xcb'); sos = raw.index(b'\xff\xda')
 assert raw[sof+4] == bits and raw[sos+5+components*2] == predictor and raw[sos+7+components*2] == point
 decoded = scratch / (name+'.pnm')
 if sampling:
  # Companion Huffman streams calibrate identical sample prediction independently
  # using libjpeg-turbo. The producer's own upsampler has an odd-height edge defect.
  huffman = scratch / (name+'-huffman.jpg')
  run([oracle] + options + [str(src), str(huffman)], env)
  run([turbo, '-nosmooth', '-pnm', '-outfile', str(decoded), str(huffman)])
  tokens, data = pnm(decoded)
  assert tokens[1:3] == [b'19', b'11']
  words = list(data) if bits <= 8 else [int.from_bytes(data[i:i+2], 'big') for i in range(0, len(data), 2)]
  assert len(words) == 19*11*3
  maximum = (1 << bits)-1
  rgba = bytearray()
  for i in range(0, len(words), 3):
   rgba.extend((v*255+maximum//2)//maximum for v in words[i:i+3]); rgba.append(255)
  (output / (name+'.jpg.nearest.rgba')).write_bytes(rgba)
 else:
  run([oracle, '-c', str(jpeg), str(decoded)])
  expected = [v >> point << point for v in values]
  data = bytes(expected) if bits <= 8 else b''.join(v.to_bytes(2, 'big') for v in expected)
  assert pnm(decoded)[1] == data, name
 return [jpeg.name, bits, components, predictor, point, restart]


rows = []
for bits in (8, 12, 16):
 for components in (1, 3):
  for restart in (0, 19, 38):
   encode(bits, components, 4, 0, restart, f'b{bits}-c{components}-r{restart}')
for bits in range(2, 17):
 for components in (1, 3):
  for predictor in range(1, 8):
   for point in (0, bits-1):
    restart = 19 if predictor % 2 else 38
    rows.append(encode(bits, components, predictor, point, restart,
                       f'b{bits}-c{components}-p{predictor}-t{point}-r{restart}'))
with (output / 'manifest.csv').open('w') as f:
 writer = csv.writer(f); writer.writerow(['name', 'bits', 'components', 'predictor', 'point', 'restart']); writer.writerows(rows)
subsampled = []
for bits in (8, 12, 16):
 for index, (sampling, columns) in enumerate((('1x1,2x2,1x2', 10), ('1x1,2x1,4x1', 5), ('1x1,1x2,1x4', 19))):
  for point in (0, bits-1):
   for restart in (0, columns*2):
    subsampled.append(encode(bits, 3, 4, point, restart,
                             f'sub-b{bits}-s{index}-t{point}-r{restart}', sampling))
with (output / 'subsampled.csv').open('w') as f:
 writer = csv.writer(f); writer.writerow(['name', 'bits', 'components', 'predictor', 'point', 'restart']); writer.writerows(subsampled)
files = sorted(p for p in output.iterdir() if p.suffix in ('.jpg', '.rgba'))
(output / 'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in files))
print(f'Generated {len(rows)+18} full-resolution and {len(subsampled)} subsampled JPEGs.')

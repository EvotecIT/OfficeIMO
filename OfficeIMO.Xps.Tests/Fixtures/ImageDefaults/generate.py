from pathlib import Path
from PIL import Image, TiffImagePlugin
import json, struct, hashlib

root = Path(__file__).resolve().parent
fixtures = root if root.name == 'ImageDefaults' else root / 'fixtures'
fixtures.mkdir(parents=True, exist_ok=True)
cases = []

def record(name, image, **kwargs):
    path = fixtures / name
    image.save(path, **kwargs)
    cases.append({'name': name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
                  'samples': [[4, 4, list(image.convert('RGBA').getpixel((4, 4)))],
                              [12, 4, list(image.convert('RGBA').getpixel((12, 4)))]]})

rgb = Image.new('RGB', (16, 8), (128, 64, 32))
for y in range(8):
    for x in range(8, 16): rgb.putpixel((x, y), (32, 128, 192))
rgba = rgb.convert('RGBA'); rgba.putalpha(128)
gray = Image.new('L', (16, 8), 128)
calibration = TiffImagePlugin.ImageFileDirectory_v2()
calibration[301] = tuple(round((i / 255) ** 1.0 * 65535) for i in range(256)) * 3
calibration[318] = (TiffImagePlugin.IFDRational(3457, 10000), TiffImagePlugin.IFDRational(3585, 10000))
calibration[319] = tuple(TiffImagePlugin.IFDRational(v, 10000) for v in (6800,3200,2650,6900,1500,600))
calibration[42240] = TiffImagePlugin.IFDRational(1, 1)
for mode, image in (('rgb', rgb), ('rgba', rgba), ('gray', gray)):
    record(mode+'-calibrated.tif', image, tiffinfo=calibration, compression='raw')
orientation = TiffImagePlugin.ImageFileDirectory_v2(); orientation[274] = 6
record('rgb-oriented.tif', rgb, tiffinfo=orientation, compression='raw')
repo = next(p for p in root.parents if (p / 'OfficeIMO.Core').is_dir())
profile = repo / 'OfficeIMO.Drawing.Tests/TestAssets/IccColorCorpus/littlecms-rgb-matrix.icc'
record('rgb-oriented-profile.tif', rgb, tiffinfo=orientation, compression='raw', icc_profile=profile.read_bytes())
record('rgba-unspecified.tif', rgba, compression='raw')
path = fixtures / 'rgba-unspecified.tif'; data = bytearray(path.read_bytes())
ifd = struct.unpack_from('<I', data, 4)[0]
for i in range(struct.unpack_from('<H', data, ifd)[0]):
    entry = ifd + 2 + i * 12
    if struct.unpack_from('<H', data, entry)[0] == 338: struct.pack_into('<H', data, entry + 8, 0)
path.write_bytes(data); cases[-1]['sha256'] = hashlib.sha256(data).hexdigest()
for sample in cases[-1]['samples']: sample[2][3] = 255
exif = Image.Exif(); exif[318] = calibration[318]; exif[319] = calibration[319]; exif[42240] = TiffImagePlugin.IFDRational(1, 1)
for mode, image in (('rgb', rgb), ('gray', gray)):
    record(mode+'-calibrated.jpg', image, exif=exif, quality=100, subsampling=0)
    decoded = Image.open(fixtures / cases[-1]['name']).convert('RGBA')
    for sample in cases[-1]['samples']: sample[2] = list(decoded.getpixel(tuple(sample[:2])))
(fixtures / 'manifest.json').write_text(json.dumps(cases, indent=2)+'\n')
print(len(cases), 'independently encoded JPEG/TIFF image fixtures')

rows=['file,x,y,r,g,b,a']
for case in cases:
    for x,y,color in case['samples']: rows.append(','.join(map(str,[case['name'],x,y]+color)))
(fixtures/'expected.csv').write_text('\n'.join(rows)+'\n')

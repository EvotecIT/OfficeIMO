#!/usr/bin/env python3
"""Opt-in SVT-AV1 4.0.1 producer for a reduced-header, tiled mixed-skip still frame."""
import argparse,hashlib,json,pathlib,subprocess,tempfile

EXPECTED='a7df0ead410eeec2758c51d0764f73463aacbccced234f18dcd97af5ab74c46f'

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--output',required=True,type=pathlib.Path)
    ap.add_argument('--ffmpeg',default='ffmpeg');args=ap.parse_args()
    output=args.output.resolve();output.parent.mkdir(parents=True,exist_ok=True)
    version=subprocess.check_output([args.ffmpeg,'-version'],text=True).splitlines()[0]
    if not version.startswith('ffmpeg version 8.0.1 '):raise ValueError('Expected FFmpeg 8.0.1')
    pixels=bytearray()
    for p in range(3):
        sub=bool(p)
        for y in range(136>>sub):
            for x in range(256>>sub):
                pixels.append(128 if x<(128>>sub) and y<(96>>sub) else 48+(x*3+y*2+p*17)%127+((x//13+y//11)%2)*31)
    with tempfile.TemporaryDirectory(prefix='svt-cdef-',dir=output.parent) as scratch:
        raw=pathlib.Path(scratch)/'input.yuv';encoded=pathlib.Path(scratch)/'frame.obu';raw.write_bytes(pixels)
        command=[args.ffmpeg,'-hide_banner','-nostdin','-loglevel','info','-f','rawvideo','-pixel_format','yuv420p',
            '-video_size','256x136','-framerate','1','-i',str(raw),'-frames:v','1','-c:v','libsvtav1','-preset','4','-crf','36',
            '-svtav1-params','avif=1:tile-columns=1:tile-rows=1:lp=1:enable-restoration=0','-f','obu',str(encoded)]
        result=subprocess.run(command,check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True)
        if 'SVT-AV1 Encoder Lib v4.0.1' not in result.stderr:raise ValueError('Expected SVT-AV1 4.0.1')
        data=encoded.read_bytes();sha=hashlib.sha256(data).hexdigest()
        if sha!=EXPECTED:raise ValueError('Producer output differs from the qualified native control')
        output.write_bytes(data)
    receipt={'ffmpeg':version,'encoder':'SVT-AV1 4.0.1','pixelsSha256':hashlib.sha256(pixels).hexdigest(),
        'inputBytes':len(pixels),'outputBytes':len(data),'outputSha256':sha,'scriptSha256':hashlib.sha256(pathlib.Path(__file__).read_bytes()).hexdigest(),
        'parameters':['<input.yuv>' if v==str(raw) else v for v in command[1:-1]],
        'boundary':'Opt-in producer only; normal tests consume the checked-in OBU and native pixel fixture.'}
    output.with_suffix('.producer.json').write_text(json.dumps(receipt,indent=2)+'\n')
    output.with_suffix('.encode.log').write_text(result.stderr)
    print(json.dumps(receipt,indent=2))

if __name__=='__main__':main()

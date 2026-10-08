"""Produce finite offline inputs before measuring renderers. Requires Pillow with AVIF.

Run from any directory. All fixture artwork and HTML are OfficeIMO-authored MIT
inputs; the embedded Baseline Sans font retains its separate SIL OFL notice.
"""
from pathlib import Path
import base64, hashlib, json
from PIL import Image, features
import PIL

ROOT = Path(__file__).resolve().parents[5]
OUT = Path(__file__).resolve().parent
FONT = ROOT / 'OfficeIMO.TestAssets/Fonts/OfficeIMOBaselineSans-Regular.ttf'
ALPHA = ROOT / 'OfficeIMO.Drawing.Tests/TestAssets/WebpAlpha'

def sha(data):
    return hashlib.sha256(data).hexdigest()

def uri(mime, data):
    return 'data:' + mime + ';base64,' + base64.b64encode(data).decode('ascii')

font_bytes = FONT.read_bytes()
font_css = "@font-face{font-family:FixtureSans;src:url('" + uri('font/ttf', font_bytes) + "')}"
css = font_css + "@page{size:A4;margin:24px}body{margin:0;font:16px/24px FixtureSans;color:#111}h1{font:24px/32px FixtureSans;margin:0 0 16px}p{margin:0 0 8px}img{vertical-align:top}.swatch{display:inline-block;padding:8px;margin:4px;border:1px solid #888}"
cases = []

def case(id, group, body, extra='', markers=None, criteria=None, oracle='Chromium print and published managed reference', images=0, links=None):
    text = '<!doctype html><html lang="en-US"><meta charset="utf-8"><title>' + id + '</title><style>' + css + extra + '</style><body><h1>' + id + '</h1>' + body + '<p>END-' + id + '</p></body></html>\n'
    data = text.encode('utf-8')
    (OUT / (id + '.html')).write_bytes(data)
    cases.append({'id':id,'group':group,'path':id+'.html','length':len(data),'sha256':sha(data),'encoding':'utf-8','textMarkers':[id]+(markers or [])+['END-'+id], 'requiredLinks':links or [], 'minimumImageOccurrences':images, 'criteria':criteria or [],'oracle':oracle,'sourceLicense':'MIT','timeoutSeconds':30})

def paragraphs(prefix, count):
    return ''.join('<p>'+prefix+'-'+str(i).zfill(2)+' retained paragraph text.</p>' for i in range(count))

def alpha_case(kind):
    body=''
    for filter in range(4):
        data=(ALPHA / (kind+'-filter-'+str(filter)+'.webp')).read_bytes()
        for color in ('#fff','#111'):
            body+='<div class="swatch" style="background:'+color+'"><img width="98" height="66" alt="filter-'+str(filter)+'" src="'+uri('image/webp',data)+'"></div>'
        body+='<p>FILTER-'+str(filter)+'</p>'
    case('webp-'+kind+'-filters','webp-alpha',body,markers=['FILTER-'+str(i) for i in range(4)],criteria=['All four filters on white/dark backgrounds; compare color/alpha regions to independently decoded .rgba fixtures.','No decode-fallback or omitted-image diagnostic; each of eight image instances retained.'],images=8)

alpha_case('raw')
alpha_case('compressed')

assets=[]
for alpha in (False,True):
    name='avif-alpha' if alpha else 'avif-opaque'
    image=Image.new('RGBA' if alpha else 'RGB',(49,33))
    for y in range(33):
        for x in range(49):
            rgb=(220 if x<24 else 35,45 if y<16 else 205,60 if x<24 else 210)
            image.putpixel((x,y),rgb+((x*255//48,) if alpha else ()))
    path=OUT/(name+'.avif')
    image.save(path,format='AVIF',quality=90,speed=6,subsampling='4:2:0',bit_depth=8)
    decoded=Image.open(path).convert('RGBA'); decoded.load()
    rgba=decoded.tobytes(); (OUT/(name+'.rgba')).write_bytes(rgba)
    assets.append({'path':path.name,'sha256':sha(path.read_bytes()),'length':path.stat().st_size,'referencePixels':name+'.rgba','referenceSha256':sha(rgba),'width':49,'height':33,'producer':'Pillow '+PIL.__version__+'/libavif '+str(features.version('avif')),'profile':'8-bit YUV420 still with alpha auxiliary image' if alpha else '8-bit YUV420 still'})
    src=uri('image/avif',path.read_bytes())
    body=''.join('<div class="swatch" style="background:'+color+'"><img width="196" height="132" alt="AVIF-'+('ALPHA' if alpha else 'OPAQUE')+'" src="'+src+'"></div>' for color in ('#fff','#111'))
    case(name,'avif-still',body,criteria=['Two image instances retained; contrast backgrounds and channel values agree with independently decoded .rgba (RGB tolerance 3, alpha exact).','Encoded/pixel bounds and malformed cases must be qualified through Core before support promotion.'],images=2)

for placement in ('top','bottom','snap'):
    body=paragraphs('BEFORE',15)+'<div class="float">FLOAT-'+placement.upper()+'</div>'+paragraphs('AFTER',35)
    case('page-float-'+placement,'page-floats',body,'.float{float:'+placement+';float-reference:page;background:#fee080;height:72px;border:2px solid #800}',markers=['BEFORE-'+str(i).zfill(2) for i in range(15)]+['FLOAT-'+placement.upper()]+['AFTER-'+str(i).zfill(2) for i in range(35)],criteria=['All unique paragraph/float markers exactly once; no float/body overlap.','Top/bottom occupy the corresponding page content edge. Snap chooses nearest edge under the explicitly labeled reference contract.'],oracle='Pinned reference plus CSS Page Floats for top/bottom; snap is reference-defined compatibility, not a standards claim')

for placement in ('left','top','bottom'):
    body='<div class="columns">'+paragraphs('COLBEFORE',4)+'<div class="float">COLUMN-'+placement.upper()+'</div>'+paragraphs('COLAFTER',38)+'<a href="https://example.test/column-link">COLUMN-LINK</a></div>'
    case('column-float-'+placement,'column-floats',body,'.columns{column-count:2;column-gap:24px;height:720px;column-fill:auto}.float{float:'+placement+';float-reference:column;background:#cde;width:120px;height:96px}',markers=['COLUMN-'+placement.upper(),'COLBEFORE-00','COLAFTER-37','COLUMN-LINK'],criteria=['Float consumes space only in originating column; neighboring column retains its full content band.','All paragraph markers once; continuation has no overlap or loss; link retained.'],oracle='Pinned reference and column geometry; Chromium only supplies supported inline float controls',links=['https://example.test/column-link'])

notes='<div class="columns"><p>CALL-A <span class="note">NOTE-A short originating-column note.</span> body.</p>'+paragraphs('NOTE-BODY',26)+'<p>CALL-B <span class="note">'+''.join('LONG-NOTE-'+str(i).zfill(2)+' retained note text.<br>' for i in range(35))+'</span> body.</p>'+paragraphs('NOTE-TAIL',28)+'</div>'
case('column-notes','column-notes',notes,'.columns{column-count:2;column-gap:24px;height:720px;column-fill:auto}.note{float:footnote;float-reference:column;font-size:12px;line-height:18px}',markers=['CALL-A','NOTE-A','CALL-B','LONG-NOTE-00','LONG-NOTE-34','NOTE-TAIL-27'],criteria=['Each note bound to column of its call; all 35 long-note markers once and in order; continuation reserves space without body overlap.'],oracle='Pinned reference plus GCPM footnote semantics; Chromium does not implement this contract')

hyphen='<div class="words" lang="en-US">representation extraordinary internationalization <p lang="de-DE">Silbentrennung Donaudampfschifffahrt außergewöhnlich</p><p lang="zz">unsupportedlanguage</p></div><div class="manual" lang="en-US">represen&shy;tation</div><div class="none">representation</div>'
case('language-hyphenation','language-hyphenation',hyphen,'.words,.manual,.none{width:105px;border:1px solid #555;hyphenate-limit-chars:5 2 2}.words{hyphens:auto}.manual{hyphens:manual}.none{hyphens:none}',criteria=['en-US/de-DE use declared language patterns; nested lang overrides inherited language.','Manual/none controls; unsupported language has explicit fallback, no invented pattern.','Original words preserved semantically; dictionary license/provenance recorded independently.'],oracle='Pinned reference and browser with recorded dictionary/platform availability; approved pattern corpus is decisive')

svg_head='<svg xmlns="http://www.w3.org/2000/svg" width="360" height="180" viewBox="0 0 360 180">'
patterns=svg_head+'<defs><pattern id="u" patternUnits="userSpaceOnUse" width="20" height="20"><rect width="20" height="20" fill="#f00"/><rect width="10" height="10" fill="#00f"/></pattern><pattern id="b" width=".2" height=".25" patternContentUnits="objectBoundingBox" patternTransform="rotate(20)"><rect width=".2" height=".25" fill="#ff0"/><rect width=".1" height=".125" fill="#080"/></pattern><clipPath id="clip"><circle cx="270" cy="90" r="65"/></clipPath></defs><rect x="10" y="10" width="150" height="160" fill="url(#u)"/><rect x="180" y="10" width="170" height="160" fill="url(#b)" clip-path="url(#clip)"/></svg>'
case('svg-patterns','svg-patterns',patterns,criteria=['Left grid repeats 20px red/blue tiles; right rotated object-bounding-box pattern stays inside circular clip.','Independent color-region/clip checks; no fallback-colored silhouette.'])
filters=svg_head+'<defs><filter id="shadow" x="-30%" y="-30%" width="180%" height="180%"><feGaussianBlur in="SourceAlpha" stdDeviation="4" result="blur"/><feOffset in="blur" dx="12" dy="8" result="off"/><feComposite in="SourceGraphic" in2="off" operator="over"/></filter><filter id="color"><feColorMatrix type="matrix" values="0 0 1 0 0 0 1 0 0 0 1 0 0 0 0 0 0 0 1 0" result="matrix"/><feBlend in="SourceGraphic" in2="matrix" mode="multiply"/></filter></defs><rect x="35" y="35" width="100" height="90" fill="#e03020" filter="url(#shadow)"/><rect x="210" y="35" width="100" height="90" fill="#c040a0" filter="url(#color)"/></svg>'
case('svg-filter-graph','svg-filter-graph',filters,criteria=['Blur/offset/source-alpha chain paints visible shadow beyond source bounds without clipping.','Color matrix/blend produces independent reference color; graph inputs/results honored.'])
svg_resources=svg_head+'<defs><linearGradient id="a" href="#b"/><linearGradient id="b" href="#a"/><symbol id="tile"><rect width="50" height="40" fill="#209040"/></symbol></defs><use href="#tile" x="20" y="20"/><rect x="130" y="20" width="70" height="70" fill="url(#a) red"/><a href="https://example.test/svg-link"><rect x="240" y="20" width="80" height="80" fill="#2040c0"/></a></svg>'
case('svg-resources','svg-resources',svg_resources,criteria=['Local symbol retained in green; cyclic gradient terminates within work limit with explicit fallback/diagnostic.','SVG link retained; the denied-resource sibling must enforce offline resolution.'],links=['https://example.test/svg-link'])

denied=svg_head+'<rect x="10" y="10" width="60" height="60" fill="#209040"/><image x="100" y="10" width="60" height="60" href="file:///officeimo-qualification-denied.png"/></svg>'
case('svg-resource-denied','svg-resources',denied+'<p>DENIED-RESOURCE-SAFE-TEXT</p>',markers=['DENIED-RESOURCE-SAFE-TEXT'],criteria=['File URI is rejected without accessing the filesystem; adjacent green geometry and safe text remain.','Rejection/fallback is observable in diagnostics; resource/work limits bound failed resolution.'])

math='<p>FORMULA-PREFIX <math xmlns="http://www.w3.org/1998/Math/MathML" alttext="x squared plus y over two"><mfrac><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><mi>y</mi></mrow><mn>2</mn></mfrac></math> FORMULA-SUFFIX</p><a href="https://example.test/math-link">MATH-LINK</a>'
for id,group in [('mathml-semantics','mathml-semantics'),('pdf-conformance','pdf-conformance')]:
    case(id,group,math,markers=['FORMULA-PREFIX','FORMULA-SUFFIX','MATH-LINK'],criteria=['Formula vector fraction/superscript geometry retained.','Tagged Formula has accessible description and exact original MathML associated-file mapping.','Conformance case requests PDF/A-3b and PDF/UA-1 separately through each engine supported public policy API; validate exact bytes with pinned independent rule sets.'] if group=='pdf-conformance' else ['Formula vector geometry, Formula structure/description, exact original MathML associated-file mapping.'],oracle='Browser appearance plus independent PDF object inspection; veraPDF policy validators for requested archival/accessibility profiles',links=['https://example.test/math-link'])

manifest={'schema':'officeimo.html.static-gap-corpus','version':1,'declaredSourceHead':'802cd545cd866d667f5050fcb55b8baa664e8b76','reference':{'package':'PeachPDF','version':'0.9.20','license':'BSD-3-Clause'},'font':{'path':str(FONT.relative_to(ROOT)),'sha256':sha(font_bytes),'family':'OfficeIMO Baseline Sans','license':'SIL OFL 1.1','notice':'OfficeIMO.TestAssets/Fonts/OFL-Carlito.txt','embeddedInEveryInput':True},'inputPolicy':{'resources':'Self-contained data URIs and local SVG references; outbound network denied. Hyperlink targets are annotations, never fetched.','page':'A4 portrait, authored 24 CSS px margins, zero caller margins, print media.','text':'Every unique visible marker exactly once after independent extraction; full per-case text also checked for omission/duplication.','visual':'Compare all pages at 96 dpi. Inspect declared geometry/color/clip/overlap criteria, not page-count equality.','diagnostics':'Retain all diagnostics and classify omissions/fallbacks against authored capability.','performance':'30-second per-case execution cap; existing H4/NASA operation ceilings unchanged.'},'assets':assets,'cases':cases}
(OUT/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
print(json.dumps({'cases':len(cases),'htmlBytes':sum(c['length'] for c in cases),'manifestSha256':sha((OUT/'manifest.json').read_bytes()),'avifProducer':assets[0]['producer']}))

import hashlib
import json
import pathlib
import subprocess
import reportlab
from reportlab.pdfgen.canvas import Canvas

root = pathlib.Path(__file__).resolve().parent
root.mkdir(exist_ok=True)
cases = []

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def start(name):
    canvas = Canvas(str(root / (name + '-native.pdf')), pagesize=(612, 792), invariant=True, pageCompression=0)
    canvas.setFont('Helvetica', 12)
    return canvas

def finish(name, canvas, reading, table):
    canvas.showPage()
    canvas.save()
    native = root / (name + '-native.pdf')
    subprocess.run(['pdftoppm', '-r', '150', '-singlefile', '-png', str(native), str(root / (name + '-scan'))], check=True)
    png = root / (name + '-scan.png')
    scan = Canvas(str(root / (name + '-scan.pdf')), pagesize=(612, 792), invariant=True)
    scan.drawImage(str(png), 0, 0, width=612, height=792)
    scan.showPage()
    scan.save()
    subprocess.run(['tesseract', str(png), str(root / name), '-l', 'eng', '--psm', '3', '--dpi', '150',
                    '-c', 'tessedit_create_tsv=1'], check=True)
    (root / (name + '-provider.tsv')).write_bytes((root / (name + '.tsv')).read_bytes())
    (root / (name + '.tsv')).unlink()
    cases.append(dict(id=name, languages='eng', clockwiseDegrees=0, rightToLeft=False, readingOrder=reading,
                      table=table, caption='', acceptance=dict(maximumCharacterErrorRate=0,
                          maximumWordErrorRate=0, requireCompleteReadingOrder=True, requireExactTables=True),
                      files={suffix: digest(root / (name + '-' + suffix))
                          for suffix in ['native.pdf', 'scan.pdf', 'scan.png', 'provider.tsv']}))

name = 'english-columns'
canvas = start(name)
canvas.setFont('Helvetica-Bold', 18)
canvas.drawString(44, 748, 'Quarterly Operations Review')
canvas.setFont('Helvetica', 12)
left = ['Northern region opened in April.', 'The first shipment arrived on Monday.', 'Sales increased by twelve percent.', 'The audit found no missing records.']
right = ['Southern region opened in June.', 'The second shipment arrived on Friday.', 'Returns decreased by seven percent.', 'The review requires two signatures.']
for i, (a, b) in enumerate(zip(left, right)):
    canvas.drawString(44, 690-i*25, a)
    canvas.drawString(320, 690-i*25, b)
canvas.drawString(44, 495, 'Report complete.')
finish(name, canvas, ['Quarterly Operations Review'] + left + right + ['Report complete.'], [])

name = 'english-ledger'
canvas = start(name)
canvas.setFont('Helvetica-Bold', 18)
canvas.drawString(44, 748, 'Reviewed Service Ledger')
canvas.setFont('Helvetica', 12)
table = [['Description', 'Units', 'Rate', 'Amount'], ['Initial service', '12', '9.99', '119.88'],
         ['Return credit', '3', '-4.50', '-13.50'], ['Monthly plan', '1', '89.00', '89.00']]
xs = [44, 295, 379, 483]
for i, row in enumerate(table):
    for x, value in zip(xs, row):
        canvas.drawString(x, 680-i*35, value)
canvas.drawString(44, 460, 'Report complete.')
finish(name, canvas, ['Reviewed Service Ledger'] + [' '.join(row) for row in table] + ['Report complete.'], table)
(root / 'manifest.json').write_text(json.dumps(dict(producer='ReportLab', reportlab=reportlab.Version,
    rasterProducer=subprocess.check_output(['pdftoppm', '-v'], stderr=subprocess.STDOUT, text=True).splitlines()[0],
    recordedProvider=subprocess.check_output(['tesseract', '--version'], text=True).splitlines()[0],
    dpi=150, recordedProviderDpi=150, cases=cases), indent=2) + '\n')

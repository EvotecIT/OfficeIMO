"""Independent ReportLab/pypdf AcroForms for bounded form-OCR acceptance.

This validation-only producer is outside the OfficeIMO runtime/package graph.
"""
from pathlib import Path
import hashlib
import json
import subprocess
import shutil

import pypdf
import reportlab
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    ArrayObject, BooleanObject, ByteStringObject, DecodedStreamObject, DictionaryObject,
    NameObject, NumberObject, RectangleObject, TextStringObject,
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

import argparse
parser = argparse.ArgumentParser()
parser.add_argument("output", type=Path, help="Explicit disposable output directory")
ROOT = parser.parse_args().output.resolve()
ROOT.mkdir(parents=True, exist_ok=True)
POPPLER = shutil.which("pdftoppm")
if POPPLER is None:
    raise RuntimeError("Install validation-only Poppler and put pdftoppm on PATH.")
WIDTH, HEIGHT = letter
FIELDS = [
    dict(name="FullName", page=1, x=110, y=652, w=330, h=28,
         value="Alex Morgan", maxlen=40, label="Full name", kind="text"),
    dict(name="SerialCode", page=1, x=110, y=570, w=216, h=28,
         value="A1B2C3", maxlen=6, label="Six-character code", kind="comb"),
    dict(name="Country", page=1, x=110, y=488, w=230, h=28,
         value="PL", display="Poland", label="Country", kind="choice"),
    dict(name="Amount", page=1, x=110, y=406, w=230, h=28,
         value="123.45", maxlen=12, label="Amount (0 to 1000, two decimals)", kind="number"),
    dict(name="Reference", page=2, x=110, y=652, w=330, h=28,
         value="INV-2048", maxlen=16, label="Reference", kind="text"),
    dict(name="ReadOnly", page=2, x=110, y=570, w=230, h=28,
         value="KEEP", maxlen=16, label="Read-only value", kind="readonly"),
    dict(name="CustomRule", page=2, x=110, y=488, w=230, h=28,
         value="77", maxlen=16, label="Unsupported custom rule", kind="unsupported"),
]


def make_reportlab(path, background=False):
    pdf = canvas.Canvas(str(path), pagesize=letter, invariant=1)
    pdf.setTitle("Independent form-OCR producer fixture")
    pdf.setAuthor("OfficeIMO validation fixture producer")
    for page in (1, 2):
        if background:
            pdf.drawImage(str(ROOT / f"scan-{page}.png"), 0, 0,
                          width=WIDTH, height=HEIGHT)
        else:
            pdf.setFillColor(colors.HexColor("#17324d"))
            pdf.setFont("Helvetica-Bold", 18)
            pdf.drawString(60, 740, "Form value review")
            pdf.setFont("Helvetica", 10)
            pdf.drawString(60, 716, "Independent ReportLab form and raster source")
            pdf.drawString(60, 52, f"Page {page} of 2")
        for field in [item for item in FIELDS if item["page"] == page]:
            if not background:
                pdf.setFillColor(colors.black)
                pdf.setFont("Helvetica", 11)
                pdf.drawString(field["x"], field["y"] + 40, field["label"])
            common = dict(name=field["name"], tooltip=field["label"],
                          x=field["x"], y=field["y"], width=field["w"],
                          height=field["h"], fontName="Helvetica", fontSize=13,
                          borderColor=colors.HexColor("#708499"),
                          fillColor=colors.white, textColor=colors.black,
                          forceBorder=False)
            if field["kind"] == "choice":
                # ReportLab accepts (display label, export value) tuples.
                pdf.acroForm.choice(value="PL", options=[("Poland", "PL"),
                    ("Germany", "DE"), ("France", "FR")], **common)
            else:
                flags = "comb" if field["kind"] == "comb" else "readOnly" if field["kind"] == "readonly" else ""
                value = field["value"] if not background or field["kind"] == "readonly" else ""
                pdf.acroForm.textfield(value=value, maxlen=field["maxlen"],
                                      fieldFlags=flags, **common)
        pdf.showPage()
    pdf.save()


def js(source):
    return DictionaryObject({NameObject("/S"): NameObject("/JavaScript"),
                             NameObject("/JS"): TextStringObject(source)})


def parentize_scanned(source, target):
    writer = PdfWriter()
    writer.clone_document_from_reader(PdfReader(source))
    acro = writer._root_object["/AcroForm"].get_object()
    by_name = {item["name"]: item for item in FIELDS}
    parents = ArrayObject()
    # Put actions on real parent field dictionaries, rather than widgets.
    field_keys = ("/FT", "/T", "/TU", "/Ff", "/V", "/DV", "/DA", "/MaxLen", "/Opt", "/I", "/AA")
    for ref in acro["/Fields"]:
        widget = ref.get_object()
        name = str(widget["/T"])
        expected = by_name[name]
        parent = DictionaryObject({NameObject(k): widget[k] for k in field_keys if k in widget})
        parent[NameObject("/Kids")] = ArrayObject([ref])
        parent[NameObject("/V")] = TextStringObject(expected["value"] if expected["kind"] == "readonly" else "")
        parent[NameObject("/DV")] = parent["/V"]
        parent.pop(NameObject("/I"), None)
        if name == "Amount":
            parent[NameObject("/AA")] = DictionaryObject({
                NameObject("/K"): js("AFNumber_Keystroke(2, 0, 0, 0, '', true);"),
                NameObject("/F"): js("AFNumber_Format(2, 0, 0, 0, '', true);"),
                NameObject("/V"): js("AFRange_Validate(true, 0, true, 1000);"),
            })
        elif name == "CustomRule":
            parent[NameObject("/AA")] = DictionaryObject({
                NameObject("/V"): js("event.value = eval('77');")})
        parent_ref = writer._add_object(parent)
        parents.append(parent_ref)
        for key in field_keys:
            widget.pop(NameObject(key), None)
        widget[NameObject("/Parent")] = parent_ref
        widget.pop(NameObject("/MK"), None)
        # Empty normal appearance remains transparent over scanned values.
        # Logical /V stays empty; this is scanned evidence, not filled data.
        appearance = DecodedStreamObject()
        appearance.set_data(f"q 0.4 0.5 0.6 RG 0.5 w 0 0 {expected['w']} {expected['h']} re S Q\n".encode())
        appearance.update({NameObject("/Type"): NameObject("/XObject"),
            NameObject("/Subtype"): NameObject("/Form"),
            NameObject("/BBox"): RectangleObject([0, 0, expected["w"], expected["h"]]),
            NameObject("/Resources"): DictionaryObject()})
        widget[NameObject("/AP")] = DictionaryObject({NameObject("/N"): writer._add_object(appearance)})
    acro[NameObject("/Fields")] = parents
    acro[NameObject("/NeedAppearances")] = BooleanObject(False)
    writer.write(target)


def add_test_certification(source, target):
    """Synthetic DocMDP policy metadata, not a cryptographically valid signature."""
    writer = PdfWriter()
    writer.clone_document_from_reader(PdfReader(source))
    signature = DictionaryObject({
        NameObject("/Type"): NameObject("/Sig"),
        NameObject("/Filter"): NameObject("/Adobe.PPKLite"),
        NameObject("/SubFilter"): NameObject("/adbe.pkcs7.detached"),
        NameObject("/ByteRange"): ArrayObject([NumberObject(n) for n in (0, 10, 20, 30)]),
        NameObject("/Contents"): ByteStringObject(b"\x00\x11\x22"),
        NameObject("/Reference"): ArrayObject([DictionaryObject({
            NameObject("/Type"): NameObject("/SigRef"),
            NameObject("/TransformMethod"): NameObject("/DocMDP"),
            NameObject("/TransformParams"): DictionaryObject({
                NameObject("/Type"): NameObject("/TransformParams"),
                NameObject("/V"): NameObject("/1.2"), NameObject("/P"): NumberObject(2),
            }),
        })]),
    })
    reference = writer._add_object(signature)
    field = writer._add_object(DictionaryObject({
        NameObject("/Type"): NameObject("/Annot"), NameObject("/Subtype"): NameObject("/Widget"),
        NameObject("/FT"): NameObject("/Sig"), NameObject("/T"): TextStringObject("Approval"),
        NameObject("/V"): reference, NameObject("/Rect"): RectangleObject([0, 0, 1, 1]),
        NameObject("/P"): writer.pages[0].indirect_reference,
    }))
    writer.pages[0]["/Annots"].append(field)
    writer._root_object["/AcroForm"]["/Fields"].append(field)
    writer._root_object["/AcroForm"][NameObject("/SigFlags")] = NumberObject(3)
    writer._root_object[NameObject("/Perms")] = DictionaryObject({NameObject("/DocMDP"): reference})
    writer.write(target)


filled = ROOT / "reportlab-form.pdf"
make_reportlab(filled)
certified = ROOT / "reportlab-certified-form.pdf"
add_test_certification(filled, certified)
subprocess.run([POPPLER, "-r", "170", "-png", str(filled), str(ROOT / "scan")], check=True)
staging = ROOT / "raster-form-staging.pdf"
make_reportlab(staging, background=True)
scanned = ROOT / "reportlab-scanned-form.pdf"
parentize_scanned(staging, scanned)
reader = PdfReader(scanned)
writer = PdfWriter()
writer.clone_document_from_reader(reader)
writer.pages[1][NameObject("/CropBox")] = RectangleObject([20, 20, 590, 770])
writer.pages[1][NameObject("/Rotate")] = NumberObject(90)
writer.pages[1][NameObject("/UserUnit")] = NumberObject(2)
rotated = ROOT / "reportlab-rotated-form.pdf"
writer.write(rotated)

files = []
for path in (filled, scanned, rotated, certified):
    check = PdfReader(path)
    logical = check.get_fields()
    expected_names = set(item["name"] for item in FIELDS)
    if path == certified:
        expected_names.add("Approval")
    assert set(logical) == expected_names
    widget_count = 0
    for page in check.pages:
        for ref in page["/Annots"]:
            widget = ref.get_object()
            assert widget["/Subtype"] == "/Widget"
            parent = widget.get("/Parent")
            field = parent.get_object() if parent else widget
            assert str(field["/V"]) == str(logical[str(field["/T"])] ["/V"])
            if field["/FT"] != "/Sig":
                assert len(widget["/AP"]["/N"].get_object().get_data()) > 0
            widget_count += 1
    assert widget_count == len(expected_names)
    files.append(dict(path=str(path), logicalBytes=path.stat().st_size,
        sha256=hashlib.sha256(path.read_bytes()).hexdigest(), pageCount=len(check.pages),
        fieldCount=len(logical), widgetCount=widget_count))
manifest = dict(producer={"reportlab": reportlab.Version, "pypdf": pypdf.__version__,
    "rendering": subprocess.check_output([POPPLER, "-v"], stderr=subprocess.STDOUT, text=True).splitlines()[0]},
    sourceScript=str(Path(__file__).resolve()), files=files, expected=FIELDS,
    constraints={"Amount": {"decimalPlaces": 2, "minimum": 0, "maximum": 1000,
        "actionsLocation": "parent field /AA, separated from widget"},
        "CustomRule": "unsupported inert validation script, proposals must be rejected",
        "ReadOnly": "must retain canonical KEEP value"},
    qualification="Independent producer/parser/raster evidence; synthetic form content, no claim of Adobe/native application or general OCR accuracy.")
(ROOT / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n")
print(json.dumps({"files": files, "producer": manifest["producer"]}, indent=2))

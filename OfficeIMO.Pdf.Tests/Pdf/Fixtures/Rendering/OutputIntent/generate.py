"""Opt-in producer/reference generation: ReportLab, pypdf, Pillow/LittleCMS and MuPDF.

Run in an isolated validation environment, never as part of a product build.
"""

import ctypes as C
import ctypes.util
import io
import json
from pathlib import Path
import subprocess
import struct
import tempfile

from PIL import Image, ImageCms
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, DecodedStreamObject, DictionaryObject, NameObject, NumberObject
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[4]


class XyY(C.Structure):
    _fields_ = [("x", C.c_double), ("y", C.c_double), ("Y", C.c_double)]


class Primaries(C.Structure):
    _fields_ = [("Red", XyY), ("Green", XyY), ("Blue", XyY)]


def profile():
    """A narrow RGB gamut makes converting before and after blending distinguishable."""
    lcms = C.CDLL(ctypes.util.find_library("lcms2"))
    lcms.cmsBuildGamma.argtypes = [C.c_void_p, C.c_double]
    lcms.cmsBuildGamma.restype = C.c_void_p
    lcms.cmsCreateRGBProfile.argtypes = [C.POINTER(XyY), C.POINTER(Primaries), C.POINTER(C.c_void_p)]
    lcms.cmsCreateRGBProfile.restype = C.c_void_p
    lcms.cmsSaveProfileToMem.argtypes = [C.c_void_p, C.c_void_p, C.POINTER(C.c_uint32)]
    lcms.cmsCloseProfile.argtypes = [C.c_void_p]
    lcms.cmsFreeToneCurve.argtypes = [C.c_void_p]
    gamma = lcms.cmsBuildGamma(None, 2.2)
    curves = (C.c_void_p * 3)(gamma, gamma, gamma)
    primaries = Primaries(XyY(.52, .33, 1), XyY(.30, .52, 1), XyY(.20, .15, 1))
    handle = lcms.cmsCreateRGBProfile(C.byref(XyY(.3127, .3290, 1)), C.byref(primaries), curves)
    size = C.c_uint32()
    assert handle and lcms.cmsSaveProfileToMem(handle, None, C.byref(size))
    data = C.create_string_buffer(size.value)
    assert lcms.cmsSaveProfileToMem(handle, data, C.byref(size))
    lcms.cmsCloseProfile(handle)
    lcms.cmsFreeToneCurve(gamma)
    result = bytearray(data.raw)
    result[24:36] = struct.pack(">6H", 2026, 1, 1, 0, 0, 0)
    return bytes(result)


def main():
    ROOT.mkdir(parents=True, exist_ok=True)
    pdfmetrics.registerFont(TTFont("Corpus", str(REPO / "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf")))
    output_profile = profile()
    plain = io.BytesIO()
    canvas = Canvas(plain, pagesize=(128, 96), invariant=1, pageCompression=0)
    labels = ["Opaque proof", "Opacity proof", "Multiply proof", "Image mask proof", ""]
    for index, label in enumerate(labels):
        canvas.setFont("Corpus", 10)
        if label:
            canvas.drawString(6, 82, label)
        canvas.setFillColorRGB(.1, .25, .9)
        canvas.rect(8, 8, 72, 56, fill=1, stroke=0)
        canvas.saveState()
        if index == 1:
            canvas.setFillAlpha(.5)
        elif index == 2:
            canvas.setBlendMode("Multiply")
        if index == 3:
            image = Image.new("RGBA", (16, 16), (230, 26, 77, 128))
            canvas.drawImage(ImageReader(image), 40, 8, 72, 56, mask="auto")
        else:
            canvas.setFillColorRGB(.9, .1, .3)
            canvas.rect(40, 8, 72, 56, fill=1, stroke=0)
        canvas.restoreState()
        canvas.showPage()
    canvas.save()
    (ROOT / "reportlab-authored.pdf").write_bytes(plain.getvalue())

    producer = PdfReader(io.BytesIO(plain.getvalue()))
    writer = PdfWriter(clone_from=producer)
    writer.pdf_header = producer.pdf_header
    writer.pages[4][NameObject("/Group")] = DictionaryObject({
        NameObject("/S"): NameObject("/Transparency"), NameObject("/CS"): NameObject("/DeviceRGB")})
    icc = DecodedStreamObject()
    icc.set_data(output_profile)
    icc[NameObject("/N")] = NumberObject(3)
    intent = DictionaryObject({NameObject("/Type"): NameObject("/OutputIntent"),
                               NameObject("/S"): NameObject("/GTS_PDFA1"),
                               NameObject("/DestOutputProfile"): writer._add_object(icc)})
    writer.root_object[NameObject("/OutputIntents")] = ArrayObject([intent])
    writer.write(ROOT / "reportlab-output-intent.pdf")

    srgb = ImageCms.createProfile("sRGB")
    destination = ImageCms.ImageCmsProfile(io.BytesIO(output_profile))
    # Retain the bounded device stage explicitly; a fused floating-point proof
    # can optimize inverse matrices away without clipping out-of-gamut RGB.
    to_device = ImageCms.buildTransform(srgb, destination, "RGB", "RGB", renderingIntent=1)
    to_display = ImageCms.buildTransform(destination, srgb, "RGB", "RGB", renderingIntent=1)
    with tempfile.TemporaryDirectory() as temporary:
        directory = Path(temporary)
        subprocess.run(["mutool", "draw", "-q", "-r", "72", "-o", str(directory / "page-%d.png"),
                        str(ROOT / "reportlab-authored.pdf")], check=True)
        samples = []
        for page in range(2, 5):
            authored = Image.open(directory / f"page-{page}.png").convert("RGB")
            reference = ImageCms.applyTransform(ImageCms.applyTransform(authored, to_device), to_display)
            reference.info.clear()  # Keep pixel baselines independent of generated ICC timestamps.
            reference.save(ROOT / f"page-{page}-reference.png")
            samples.append({"page": page, "label": labels[page - 1],
                            "pixels": [{"x": x, "y": 56, "rgb": list(reference.getpixel((x, 56)))}
                                       for x in (20, 56, 99)]})
        (ROOT / "reference.json").write_text(json.dumps(samples, indent=2) + "\n")


if __name__ == "__main__":
    main()

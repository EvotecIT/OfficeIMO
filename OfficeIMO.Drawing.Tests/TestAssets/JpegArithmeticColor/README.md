# Adobe CMYK and YCCK JPEG references

This corpus contains 64 independent libjpeg-turbo 3.2.0 encodings: eight/twelve-bit,
CMYK/YCCK, sequential/progressive, Huffman/arithmetic, quality 30/90 and three YCCK
sampling layouts (1×1, 2×1 and 2×2). Every 35×19 image has a three-MCU restart
interval and partial edge blocks. Y and K use the same sampling factors.

`manifest.csv` identifies each case. Each JPEG has nearest and high-quality native
CMYK references stored as little-endian unsigned 16-bit words, including eight-bit
samples. Native decoding disables interblock smoothing and uses the accurate integer
IDCT. Adobe samples are inverted: canonical CMYK ink levels complement all four
native channels. Pillow's eight-bit CMYK output independently confirms this polarity.

The managed decoder compares 85,120 pixel positions across both chroma modes,
with component error at most 2/255. Its approximate device-CMYK raster conversion
is checked separately; it is not a qualified default print profile.

The existing `../IccColorCorpus/littlecms-cmyk-lut.icc` supplies explicit color
management. LittleCMS 2.19 transforms full-precision native nearest samples to sRGB.
The 64 `.jpg.srgb` references use relative colorimetric intent and canonical ink
levels. The 64 eight-bit `.pdf-normal.srgb`/`.pdf-inverted.srgb` references use
absolute colorimetric intent and the named PDF Decode polarity. `icc-reference.json`
records the profile digest, reference digests and transform conventions.

Regeneration uses test-only native libraries; neither library enters production:

```sh
cc generate.c -I "$JPEG_TURBO/include" -L "$JPEG_TURBO/lib" -ljpeg -o "$TASK_SCRATCH/generate-color"
python3 generate.py "$TASK_SCRATCH/generate-color"
LCMS_LIBRARY="$LCMS_LIBRARY" python3 generate-icc.py
```

Set those paths to existing libjpeg-turbo and LittleCMS installations and an explicit
scratch directory. `SHA256SUMS` covers all 320 image/reference files. The C encoder
and Python scripts are fixture tooling, not alternate OfficeIMO runtime decoders.

Drawing tests verify native colorants, device conversion and explicit ICC output.
XPS tests preserve profiled colors through both dialects, drawing, SVG and PDF.
PDF tests independently exercise absent and inverted Decode arrays over the 32
eight-bit JPEGs, including both arithmetic scan modes. DeviceCMYK extraction and
rendering also compare absent/identity/inverted Decode arrays with implicit and
explicit equivalent color transforms, ensuring that pass-through cannot change polarity. Twelve-bit XPS export uses
normalized pixels instead of embedding a twelve-bit PDF DCT image.

Independent direct-PDF rendering covers 64 eight-bit image/Decode combinations.
MuPDF agrees within 2/255 for unsubsampled CMYK/YCCK; subsampled YCCK differences
reach 52/255. A high-quality native-chroma calibration still differs by 44/255,
so these differences are not attributed solely to a chroma-filter choice.
Ghostscript opens all cases but differences reach 245/255. The complete direct-PDF
matrix is not claimed as native pixel agreement. Managed PDF output is checked
against the native nearest-component/LittleCMS references within 5/255.

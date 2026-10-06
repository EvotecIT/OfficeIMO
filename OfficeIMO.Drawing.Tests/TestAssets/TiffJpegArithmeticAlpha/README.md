# Arithmetic JPEG TIFF color and alpha

These 288 independently produced eight-bit TIFFs use sequential arithmetic JPEG
(SOF9), three-MCU restarts and the existing TIFF alpha corpus's sample patterns.
They cover both byte orders, strips/tiles, local/shared quantization tables,
chunky/separate planes, gray polarities, RGB, CMYK and centered YCbCr at 1×1/2×2.
The extra sample is unspecified, associated alpha or unassociated alpha.
The companion [low-alpha corpus](../TiffJpegArithmeticLowAlpha/README.md) adds
192 cases with source alpha bands spanning zero through fully opaque.

LibTIFF 4.7.2 creates the containers; libjpeg-turbo 3.2.0 encodes and decodes the
arithmetic components. These remain test-only tools. Regenerate using the shared
producer, with paths to existing installations and a task scratch directory:

```sh
cc ../TiffJpegAlpha/generate.c -I "$LIBTIFF/include" -I "$JPEG_TURBO/include" -L "$LIBTIFF/lib" -L "$JPEG_TURBO/lib" -ltiff -ljpeg -o "$TASK_SCRATCH/generate-alpha"
python3 ../TiffJpegAlpha/generate.py "$TASK_SCRATCH/generate-alpha" --arithmetic
python3 ../TiffJpegAlpha/generate.py "$TASK_SCRATCH/generate-alpha" --arithmetic --low-alpha
LCMS_LIBRARY="$LCMS_LIBRARY" python3 generate-icc.py
```

The switch selects the corpus and arithmetic encoding together. An audit identifies
SOF9 in all 4,320 strips/tiles across the 480 files. `manifest.csv` describes layouts;
`.raw` contains independently decoded TIFF device samples, `.planes` retains separate
plane records, and `.rgba` applies the TIFF color/alpha contract. Empty plane files
are omitted. `SHA256SUMS` covers TIFFs and references.

The inherited producer boundaries remain explicit: YCbCr with extra samples uses
RGB container bookkeeping before patching only PhotometricInterpretation; five-channel
chunky CMYK scans are decoded independently with single-component frame headers,
without changing their entropy bytes. See the [shared producer contract](../TiffJpegAlpha/README.md).

Full-file LibTIFF decoding succeeds for 320 files: both gray polarities and RGB,
planar CMYK, and unsubsampled planar YCbCr. `decode-native.c` reads the original
TIFF strips/tiles and reconstructs device channels; all 665,000 samples match the
JPEG-derived references exactly. LibTIFF rejects the other 160 files: 40 chunky
CMYK files with a fifth component, 80 chunky YCbCr/extra-sample files and 40
subsampled planar YCbCr/alpha layouts. Component evidence is not full-file native
acceptance for those layouts.

Managed tests compare all 319,200 pixels. Ordinary alpha uses RGB tolerances 3/255,
or 6/255 for associated alpha; low-alpha tests compare visible compositing over
black and white within 3/255 and exact alpha. For all 80 CMYK cases, LittleCMS 2.19
converts native colorants through the existing explicit CMYK ICC profile, retaining
unassociated precision before conversion. `.icc-rgba` and `icc-reference.json` retain
that independent proof; managed visible compositing agrees within 3/255 with exact
alpha. TIFF colorants follow PhotometricInterpretation, without Adobe inversion.

XPS/OpenXPS tests retain alpha and require explicit ICC for CMYK. The 1,920
black/white exports cover 1,276,800 pixel-center probes per route. Independent
MuPDF PDF/SVG rendering differs by at most 4/255 and 2/255 respectively, without
warnings. GhostXPS opens all exports but differs by up to 255/255, including blank
images. The [twelve-bit corpus](../TiffJpegArithmetic12/README.md) separately qualifies
high-precision samples. Native Windows behavior and lossless arithmetic
color/alpha remain separate qualification gaps.

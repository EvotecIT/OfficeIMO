# Twelve-bit arithmetic JPEG TIFF color and alpha

These 384 35×19 TIFFs qualify the existing twelve-bit sequential arithmetic JPEG
path (SOF9). The matrix covers both byte orders, strips/tiles, local/shared tables,
chunky/separate planes, gray polarities, RGB, CMYK and centered YCbCr at 1×1/2×2.
Every layout includes ordinary opaque samples (`e-1`), an unspecified extra sample
(`e0`), associated alpha (`e1`) and unassociated alpha (`e2`). Alpha source bands
include 0, 1, 2, 3, 4, 8, 16, 32, 64, 2048, 3072, 4094 and 4095.

LibTIFF 4.7.2 creates the containers; libjpeg-turbo 3.2.0 encodes and independently
decodes twelve-bit components with three-MCU restarts. All 3,264 segments use SOF9.
The shared producer's eight-bit compatibility check regenerates 960 existing
fixtures with identical TIFF and native component/plane bytes. Tools remain
isolated fixture tooling; OfficeIMO has no new runtime dependency.

Regenerate with paths to existing native installations and task scratch storage:

```sh
cc ../TiffJpegAlpha/generate.c -I "$LIBTIFF/include" -I "$JPEG_TURBO/include" -L "$LIBTIFF/lib" -L "$JPEG_TURBO/lib" -ltiff -ljpeg -o "$TASK_SCRATCH/generate-alpha"
python3 generate.py "$TASK_SCRATCH/generate-alpha"
LCMS_LIBRARY="$LCMS_LIBRARY" python3 ../TiffJpegArithmeticAlpha/generate-icc.py
```

`.raw` stores independently decoded device words, little-endian with twelve valid
bits. `.planes` stores plane/x/y/width/height words followed by native sample words.
The Python reference reconstructs planar centered chroma with floating-point
bilinear interpolation, then applies TIFF YCbCr and alpha equations. `.rgba`
retains the projected reference; `SHA256SUMS` covers TIFF and reference bytes.
The shared producer's YCbCr metadata construction and five-component split-scan
reference boundary remain as described in the [eight-bit corpus](../TiffJpegArithmeticAlpha/README.md).

Managed tests compare all 255,360 pixels over black and white within 3/255.
Twelve-bit IDCT rounding can change alpha by one eight-bit level: a native 2048
sample and managed 2047 sample round to 128 and 127 respectively. The measured
raw discrepancy in that regression is 1/4095; this corpus permits alpha error
1/255 while retaining exact-alpha checks for the existing eight-bit corpus.
All 64 CMYK cases also use independent LittleCMS 2.19 explicit-profile references,
with unassociation at source precision before conversion.

The shared `decode-native.c` oracle compares 176 complete TIFFs exactly across
337,820 device samples. Another 96 planar/odd-width strip cases differ only in
the final column, exposing LibTIFF's twelve-bit packing limitation. LibTIFF rejects
96 five-component or YCbCr/alpha layouts. Sixteen ordinary subsampled YCbCr files
are outside this simple raw-component oracle because they require packed chroma
unit reconstruction; their native read status is recorded separately. All have
independent JPEG component references, which are not full-file native acceptance.

The 1,536 XPS/OpenXPS exports cover 1,021,440 pixel-center probes per route.
MuPDF PDF/SVG differences reach 4/255 and 2/255 respectively without warnings.
GhostXPS exits successfully on 1,408 exports, with differences up to 255/255,
and terminates with SIGSEGV on 128 CMYK strip exports. Those results do not
establish native Windows or complete native-consumer acceptance.

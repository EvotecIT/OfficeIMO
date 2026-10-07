# Fractional YCbCr conversion in JPEG-TIFF

These 60 compression-7 TIFFs cover full-resolution Huffman SOF3 and arithmetic
SOF11 YCbCr at every precision from 2 through 16 bits, in both byte orders.
They expose intermediate RGB rounding: a two-bit sample must retain fractional
color through conversion, rather than reducing the result to four RGB levels.

LibTIFF 4.7.2 wraps unchanged independently encoded component streams from
[JpegLosslessPrecision](../JpegLosslessPrecision/README.md) and
[JpegArithmeticLossless](../JpegArithmeticLossless/README.md). The manifest binds
each payload to its source SHA-256. TIFF declares full-resolution YCbCr and
explicit precision-dependent ReferenceBlackWhite values. This is constructed
color interpretation over native JPEG components, not independent whole-file
YCbCr TIFF producer/decoder acceptance.

The RGBA reference converts independently decoded/source-verified words using
TIFF's declared YCbCr equations, clips normalized RGB and rounds only at final
eight-bit output. The ICC reference supplies those fractional RGB components to
LittleCMS 2.19 with the existing DCI-P3 matrix profile and relative colorimetric
intent. The profile changes sample channels by up to 133/255 compared with the
device-color reference, so ignoring it cannot pass the comparison. Managed
output is within 1/255 of device references and 2/255 of LittleCMS references
across 11,880 pixels in each mode. The encoded payloads are opaque in this corpus;
existing 12/16-bit YCbCr-alpha fixtures additionally test conversion before
unassociation against floating-point RGB reference TIFFs.

Build `../TiffJpegArithmeticLossless/wrap.c` against test-only LibTIFF, run
`python3 generate.py /path/to/wrap`, then run `icc-reference.py` with
`LCMS_LIBRARY` pointing to the test-only LittleCMS library when library discovery
is unavailable. `SHA256SUMS` covers all 180 TIFF/reference files. Native JPEG,
LibTIFF and LittleCMS remain validation tools, not runtime dependencies.

Both XPS dialects retain device and explicit-profile paint through raster,
PNG-backed SVG and PDF readback. The 240 exports have 47,520 pixel-center probes
per route. MuPDF 1.28.2 PDF/SVG output differs by at most 2/255 without warnings.
GhostXPS 10.08.0 returns zero for 144 exports but can produce blank images, and
96 profiled cases terminate with SIGSEGV. LibTIFF's read-data check accepts only
the four Huffman 8/12-bit containers; it rejects 52 for precision and four for
an unavailable codec. Neither check establishes native whole-file color fidelity.

This corpus does not qualify low-precision alpha, subsampling, additional ICC
profiles, or wider layouts. Native Windows acceptance remains open.

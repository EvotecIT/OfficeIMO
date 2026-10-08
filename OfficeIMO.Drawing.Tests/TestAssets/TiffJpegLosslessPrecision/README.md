# Lossless JPEG-TIFF sample precision

These 192 compression-7 TIFFs qualify gray/RGB samples at 2–7, 9–11 and 13–15 bits,
complementing the existing 8/12/16-bit TIFF corpora. They cover Huffman SOF3 and
arithmetic SOF11, both TIFF byte orders, and point transforms zero and precision
minus one. All use one full-resolution chunky strip, predictor 1 and row-aligned
restarts. Huffman zero-point RGB files have separate scans; the other RGB files
use interleaved scans.

LibTIFF 4.7.2 writes the container around unchanged independently produced JPEG
payloads from [JpegLosslessPrecision](../JpegLosslessPrecision/README.md) and
[JpegArithmeticLossless](../JpegArithmeticLossless/README.md). The manifest records
the source path and SHA-256. Huffman references use the independently decoded
native sample words; arithmetic references use the source formula already checked
against the native decoder in that corpus. Projection rounds native unsigned
samples to eight-bit RGB and appends opaque alpha. `SHA256SUMS` covers the TIFFs
and reference RGBA files. This qualifies wrapping existing native JPEG streams,
not an independent whole-file TIFF encoder/decoder round trip.

Build `../TiffJpegArithmeticLossless/wrap.c` against the test-only LibTIFF library,
then run `python3 generate.py /path/to/wrap`. The optional width/height arguments
leave the older wrapper's default 19×11 output unchanged. Regeneration reproduces
all 384 fixture/reference hashes and the older 168-container corpus exactly.
Neither LibTIFF nor the native JPEG tools are runtime dependencies.

Managed validation and decoding match all 38,016 reference pixels exactly. XPS
and OpenXPS reopen and retain pixels through raster, portable PNG-backed SVG and
PDF readback. Across 384 independently rendered exports, 76,032 pixel-center
probes per route differ by at most 2/255 in MuPDF 1.28.2 PDF/SVG, without warnings.
GhostXPS 10.08.0 opens all exports but differs by up to 255/255, including blank
images. LibTIFF 4.7.2 rejects all 192 TIFFs with `Improper JPEG data precision`.
Native whole-file acceptance remains unqualified.

Other precisions' subsampled/color-profile/alpha combinations and wider layouts
are outside this corpus. DCT JPEG remains restricted to eight/twelve bits;
legacy compression-6 and non-JPEG TIFF sample-width contracts are unchanged.

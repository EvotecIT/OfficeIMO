# Subsampled lossless arithmetic JPEG TIFF

These 192 compression-7 TIFFs exercise SOF11 at eight, twelve and sixteen bits,
2×1/2×2/4×2 YCbCr subsampling, centered/cosited positioning, both byte orders,
strips and partial tiles, and contiguous/separate planes. The 35 × 19 images have
opaque, associated-alpha and straight-alpha cases. Four-component 4×2 samples use
separate TIFF planes; a single interleaved JPEG scan would exceed ten sampling
units. Each precision/sampling/position/alpha combination covers little-endian
contiguous strips, big-endian contiguous tiles, big-endian separate strips and
little-endian separate tiles where that scan layout is valid. This is not every
possible layout combination.

All seven predictors occur; point transforms are zero or two and restarts align
with two complete MCU rows. Edge strips have odd heights and partial tiles crop
the chroma grid before interpolation. Alpha uses full-resolution samples and
includes zero, low and near-opaque values.

## Independent references

Use the pinned test-only native producer and preparation script documented in
`../JpegArithmeticLossless/README.md`. The native decoder has an odd-height edge
bug even when upsampling is disabled. The generator therefore encodes companion
Huffman streams with identical inputs, prediction, sampling and restart settings,
then decodes those with libjpeg-turbo 3.2.0 through `../TiffJpegChroma16/decode.c`.
Nearest-neighbor output is checked before extracting reduced component planes.
Full-resolution luma/alpha and separately encoded planes must agree exactly with
their expected point-transformed samples.

`../tiff_chroma_reference.py` supplies independent Pillow 11.3.0 floating-point
interpolation with edge clamping. It is shared with the earlier sixteen-bit
Huffman corpus. TIFF YCbCr equations and source-precision alpha unassociation
produce the `.rgba` references. Fractional chroma and RGB survive interpolation
and unassociation until final byte projection. Tests require exact alpha and visible compositing
within 3/255 over black and white.

Compile `../TiffJpegArithmeticLosslessColor/wrap.c` against LibTIFF 4.7.2 and
`../TiffJpegChroma16/decode.c` against libjpeg-turbo, then run:

```sh
python3 generate.py /path/to/prepared/jpeg /path/to/wrap /path/to/decode /task/scratch
```

The native tools are isolated validation dependencies, not runtime or normal-build
requirements. TIFF layout and photometric declarations are constructed through
LibTIFF's raw writer; companion JPEG decoding does not qualify an independent
full-file TIFF reader. `SHA256SUMS` identifies the fixture/reference bytes.
Shared JPEG tables, unspecified extras, multi-scan chunky 4×2 alpha and native
Windows acceptance remain outside this corpus.

## Consumer evidence

The 192 TIFFs contain 2,160 arithmetic JPEG segments. Generation uses 2,754
native arithmetic/companion encoding pairs, including intermediate subsampled
planes. The system LibTIFF tool fails on 84 codec, 42 precision and 48 extra-channel
layout cases. Another 18 return zero but explicitly cannot display tile data;
these are unverified, not successful full-file decoding.

The 768 XPS/OpenXPS exports exercise 510,720 pixel-center probes per route across
black and white backgrounds. MuPDF 1.28.2 renders normalized PDF/SVG output within
4/255 and 2/255 respectively, without warnings. GhostXPS 10.08.0 opens all packages
but can render blank images (255/255 error). Native Windows and independent
full-file lossless-arithmetic TIFF acceptance remain unqualified.

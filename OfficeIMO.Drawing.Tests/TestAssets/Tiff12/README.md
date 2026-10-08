# Twelve-bit TIFF fixtures

The 228 files are independently produced and decoded with LibTIFF 4.7.2 and
libjpeg-turbo 3.2.0. They cover 35×19 images, both byte orders, chunky/separate
planes, partial strips and 16×16 edge tiles, white/black grayscale, RGB,
associated/straight RGBA, CMYK and centered 2×2 YCbCr JPEG. Unsigned packed samples use
no compression, LZW, Deflate or PackBits; JPEG uses extended sequential Huffman
coding. RGBA includes alpha 0 and 1 to exercise precision before unassociation.

`generate.c` owns native production and decoding. Build it against the two
libraries' headers and link `-ltiff -ljpeg`; pass the resulting executable to
`python3 generate.py /absolute/path/to/generate`. Neither library is a product,
normal build or correctness-test dependency. Tests consume retained fixtures.
`manifest.json` records cases; `sha256.json` records fixture and reference hashes.

Each `.tif.raw` contains little-endian native sample words. Packed TIFF references
come from LibTIFF. JPEG references decode each original compressed strip/tile
with libjpeg-turbo, including inherited quantization tables, integer slow IDCT and
RGB conversion for YCbCr. Samples remain twelve-bit until test-side final RGBA
projection. This is independent sample evidence; alpha unassociation and white-gray
inversion are checked against the declared equations rather than a native RGBA API.

Each `.tif.libtiff.raw` retains the full-file LibTIFF decode separately. LibTIFF's
[twelve-bit packing loop](https://github.com/libsdl-org/libtiff/blob/v4.7.2/libtiff/tif_jpeg.c#L1495)
only writes complete pairs. Odd-width JPEG rows can consequently omit their last
sample; 16 files differ from direct JPEG decoding. The tests use the direct JPEG
reference for those samples, not the omitted native output. The remaining 212
full-file sample outputs agree exactly with their retained references.

JPEG strips contain local tables; tiles share quantization tables and retain local
optimized Huffman tables. LibTIFF's shared-Huffman configuration emits local
redefinitions for twelve-bit encoding, contrary to the global-table restriction in
[TIFF Technical Note 2](https://libtiff.gitlab.io/libtiff/specification/technote2.html).
The generator avoids that configuration without altering encoded files afterward.

Forty CMYK cases include independent LittleCMS double-precision input references
for the existing explicit CMYK ICC profile. `generate-icc.py` records its version,
relative-colorimetric intent and reference hashes in `icc-reference.json`. This
qualifies that explicit profile, not a default SWOP profile.

This corpus does not qualify lossless JPEG, legacy table-pointer
layouts, unusual YCbCr sampling or native Windows consumption. The self-contained
legacy JPEG regression reuses a compression-7 stream under compression 6; it is
container compatibility evidence, not an independent legacy producer.

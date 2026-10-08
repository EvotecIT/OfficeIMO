# JPEG TIFF fixtures

LibTIFF 4.7.2 and its JPEG codec independently encode and decode these 160
35-by-19 images. The corpus covers both byte orders, strips and tiles, grayscale
polarities, RGB, CMYK and YCbCr, chunky and separate grayscale/RGB/CMYK planes,
and four table modes: local tables, shared quantization, shared Huffman, and both.
YCbCr covers centered 1-by-1 and 2-by-2 subsampling in chunky storage.

`generate.c` writes the TIFF and independently decodes it into the adjacent
`.tif.raw` file. Raw files contain interleaved device samples; YCbCr references
are converted to RGB by LibTIFF's JPEG codec. Tests compare every output pixel
against these references with a 3/255 channel tolerance for decoder rounding.
CMYK references use the Core device-CMYK projection; XPS requires an explicit ICC
profile and separately checks raster/SVG/PDF preservation.

Compile with local LibTIFF headers and library. Invoke:
`generate <file> <photometric> <big-endian:0|1> <tiled:0|1> <tables:0..3> <planar:1|2> <subsampling:1|2>`.
`manifest.csv` describes the matrix and `SHA256SUMS` identifies all TIFF and raw
reference bytes. LibTIFF is independent test tooling, not a product dependency.

The adapter follows [TIFF Technical Note 2](https://libtiff.gitlab.io/libtiff/specification/technote2.html).
This corpus does not qualify legacy compression 6, non-baseline JPEG processes,
extra channels, cosited chroma, separate YCbCr planes, or native Windows rendering.

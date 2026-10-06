# Sixteen-bit Huffman lossless JPEG in TIFF

These 280 fixtures use libjpeg-turbo 3.2.0 to encode and independently decode
sixteen-bit SOF3 component samples. LibTIFF 4.7.2 writes the TIFF wrappers through
its raw segment API. Its full-file decoder rejects these streams with
“Improper JPEG data precision”; this corpus does not claim independent full-file
TIFF decoder acceptance.

The 35 by 19 pixel matrix covers all seven predictors, point transforms 0/1/8/15,
separate/interleaved scans, no restart or two-row restart intervals, both byte
orders, strips/tiles and chunky/planar storage. Color models include both grayscale
polarities, RGB, CMYK and YCbCr. Gray/RGB/YCbCr include associated or unassociated
alpha, including source words 0, 1, 2, 3, 4, 8, 16, 128, 256, 1024, 32768 and 65535.
Point transforms clear discarded low bits. The independent decoder checks all
reconstructed words against the original source after that operation.

`.tif.raw` contains independently decoded little-endian page components;
`.rgba` is their color/alpha projection. `.reference.tif` stores the same native
samples uncompressed for ICC comparison. YCbCr references use normalized
floating-point RGB so fractional color survives conversion and unassociation. `.tif.jpg`
and `.tif.jpg.raw` retain the first segment and its independently decoded words.
The tests compare 239,904 JPEG component samples and 186,200 TIFF pixels exactly,
including both raw output byte orders and ICC output against reference TIFFs.

Compile `generate.c` against the installed libjpeg-turbo and LibTIFF libraries,
then run `python3 generate.py <compiled-generator>`. `SHA256SUMS` records the
fixture/reference bytes. These libraries are validation tools only.

`generate-edges.py <djpeg>` authors sixteen constant-sample JPEGs from T.81 rules
and verifies their sixteen-bit PPM output independently. They cover horizontal
and vertical subsampling, odd/even dimensions and separate/interleaved scans;
twelve compatible TIFF wrappers exercise the same samples through TIFF. These
are independent-decoder checks, not independent-producer evidence.

The same generator accepts optional precision and output-folder arguments; see
the [twelve-bit lossless corpus](../TiffJpegLossless12/README.md). Arithmetic JPEG,
other TIFF sample precisions, independent full-file sixteen-bit TIFF decoding
and native Windows acceptance remain open.

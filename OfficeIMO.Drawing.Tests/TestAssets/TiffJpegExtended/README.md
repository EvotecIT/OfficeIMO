# Extended sequential JPEG TIFF fixtures

LibTIFF 4.7.2 with libjpeg-turbo 3.2.0 independently encodes and decodes these
160 35-by-19 TIFF images. JPEG quality 1 produces eight-bit SOF1 frames with
16-bit quantization tables. The matrix covers both byte orders, strips/tiles,
grayscale polarities, RGB, CMYK, centered YCbCr, separate gray/RGB/CMYK planes,
and local/shared quantization and Huffman tables.

Compile `generate.c` with LibTIFF headers and libraries, then run
`python3 generate.py <compiled-generator>`. Adjacent `.tif.raw` files contain
independently decoded interleaved samples; YCbCr is converted to RGB by LibTIFF.
The manifest matches the baseline corpus dimensions and color contracts.
Tests compare all 106,400 pixels within 3/255 for decoder rounding.
`SHA256SUMS` covers TIFF and raw reference bytes.

The encoder's coarse-quantization warnings are expected: these files deliberately
exercise extended sequential JPEG. LibTIFF/libjpeg-turbo are isolated fixture
tools, not runtime requirements. This corpus does not qualify 12-bit samples,
lossless or arithmetic JPEG, legacy compression 6, or native Windows rendering.
TIFF continues to reject progressive and hierarchical JPEG as required by
[TIFF Technical Note 2](https://libtiff.gitlab.io/libtiff/specification/technote2.html).

`wide.c` uses libjpeg's coefficient API to write two single-block JPEGs with
quantization value 65535 and one nonzero coefficient. Their tests assert the
analytical saturated output (white for DC; four white then four black columns
for horizontal AC), guarding against fixed-point overflow. These extreme
coefficient tests are separate from LibTIFF pixel-reference qualification.

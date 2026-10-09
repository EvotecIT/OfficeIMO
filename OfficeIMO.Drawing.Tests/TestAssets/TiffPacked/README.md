# Packed TIFF decoding fixtures

These 96 synthetic fixtures are encoded and independently decoded by LibTIFF 4.7.2.
They cover 1-bit bilevel/palette and 4-bit grayscale/palette samples, both byte
orders, both grayscale polarities, strips and tiles, and uncompressed, LZW,
Deflate and PackBits data. Each image is 19 by 17 pixels. Strips contain five
rows; tiles are 16 by 16. Nonzero row/tile padding tests horizontal and vertical edge cropping.

The source value at (x,y) is `(3*x + 5*y) & ((1 << bits) - 1)`. Palette red
ramps upward, green downward, and blue uses seven times the index modulo the
palette size. The generator reads every encoded strip/tile back with LibTIFF
and checks all source samples. Tests verify Core pixels and both XPS dialects
through raster, SVG and PDF readback. `manifest.csv` identifies each case;
`SHA256SUMS` identifies the exact encoded bytes.

Compile `generate.c` with the local LibTIFF development headers/library, then
invoke `generate <file> <bits> <photometric> <big-endian:0|1> <tiled:0|1> <compression>`
for each manifest row. LibTIFF is test-only tooling and is not shipped or
required by OfficeIMO. These fixtures do not qualify CCITT, JPEG-in-TIFF,
reversed fill order, native Windows acceptance, or photographic producer input.

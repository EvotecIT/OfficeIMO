# Varying sixteen-bit JPEG chroma references

These 48 TIFF fixtures exercise sixteen-bit lossless JPEG with varying luma and
chroma, 2×1/2×2/4×2 subsampling, centered and cosited positioning, both byte orders,
chunky and separate planes, strips and partial tiles. The images are 5×3 or 17×11.

`generate.py` authors predictor-1 SOF3 streams from T.81 sample rules. The independent
libjpeg-turbo 3.2.0 decoder in `decode.c` verifies every reconstructed native word
before references are written. Its sixteen-bit output uses nearest-neighbor
upsampling even when fancy upsampling is enabled. Pillow 11.3.0's floating-point
affine interpolation therefore supplies the separate centered/cosited reference;
one replicated sample at each edge implements clamping outside the source grid.
The TIFF RGB projection uses the declared default YCbCr coefficients and reference
ranges after rounding interpolated native words.

The retained first segments provide 24 standalone JPEGs and 11,934 native component
samples per quality mode. Tests compare those samples exactly against independent
nearest decoding and Pillow interpolation, and compare 4,848 TIFF RGBA pixels
exactly. Padding beyond a partial tile's visible chroma grid is excluded from its
reference reconstruction. The 96 XPS/OpenXPS exports cover 9,696 pixel centers per
rendering route.

Compile `decode.c` against the existing libjpeg-turbo headers and library, then run:

```text
python3 generate.py <compiled-decoder>
```

`.nearest.raw` and `.bilinear.raw` hold little-endian words. `.rgba` holds eight-bit
TIFF output. `SHA256SUMS` identifies all fixture/reference bytes. Generation tools
are opt-in validation dependencies and are not required to build or run OfficeIMO.

These are specification-authored fixtures with independent decoding/interpolation
checks, not an independent producer corpus. Full-file sixteen-bit JPEG-TIFF
acceptance in another decoder and native Windows acceptance remain open.

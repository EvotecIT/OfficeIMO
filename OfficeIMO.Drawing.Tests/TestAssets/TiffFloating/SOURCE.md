# Floating TIFF fixtures

The corpus contains 304 generated image patterns encoded and decoded with libtiff
4.7.2. It covers finite IEEE half/single/double and TN3 float24 samples, both byte orders, chunky
and planar strips/tiles, uncompressed/LZW/Deflate/PackBits storage, floating-point
prediction, associated/unassociated alpha, RGB, both grayscale polarities, and
CMYK. All decoded source samples match the input exactly. RGBA references apply
normalized device-channel projection, source-precision unassociation and clipping;
unprofiled CMYK references use the existing Core approximation. XPS requires an
explicit CMYK profile and tests that conversion separately.

`provenance.json` records fixture hashes and settings. `generate.c` and `decode.c`
use libtiff only for fixture generation and comparison. `generate.py` produces the
RGB matrix, then `extra.py` adds grayscale/CMYK. Compile the C files with the local
libtiff include/library paths and `-ltiff`; run the two Python scripts in order.
The C producer requires compiler support for `_Float16`. No native tool is used
by normal tests or shipped packages. The source patterns and fixtures contain no
third-party image assets. This corpus proves the codec paths, not Windows XPS
consumer acceptance or general scientific/HDR tone mapping.

The 76 float24 fixtures combine libtiff storage/compression with unmodified
imagecodecs 2026.8.16 `imcd_float24_encode` / `imcd_float24_decode` for sample
conversion. The [TIFF Technical Note 3](http://chriscox.org/TIFFTN3d1.pdf) defines
one sign, seven exponent and sixteen fraction bits. The reference conversion
uses `0x3F0000` for 1.0 and `0x000001` for the smallest positive subnormal,
2^-78. This matches the implicit-leading-one convention used for half precision.
Hashes for the independent source and technical note are in `provenance.json`.

`generate24.c` uses the same arguments as `generate.c`, with bit depth 24. To
regenerate these fixtures, compile it with optimization enabled, libtiff and the
unmodified `imagecodecs/imcd.c` and `imcd.h` from the
[imagecodecs source](https://github.com/cgohlke/imagecodecs). Decode storage with
`decode.c`, then convert its packed samples with `imcd_float24_decode` before
comparing to the input pattern. These reference sources stay outside production
and are not needed to build or run the managed tests.

# Floating TIFF fixtures

The corpus contains 228 generated image patterns encoded and decoded with libtiff
4.7.2. It covers finite IEEE half/single/double samples, both byte orders, chunky
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

# Separate-plane JPEG TIFF fixtures

LibTIFF 4.7.2 independently encodes and decodes 48 YCbCr TIFFs with separate
JPEG planes. The 67-by-35 images cover both byte orders, strips and tiles, local
or shared tables, and horizontal/vertical chroma sampling of 1/1, 2/1, 2/2,
4/1, 4/2 and 4/4. Odd image edges and multiple tiles exercise reduced chroma
frame dimensions and padding.

`generate.c` produces the TIFFs and the `.tif.planes` reference streams. Each
reference record contains five little-endian 32-bit words (plane, x, y, width,
height), followed by the independently decoded sample bytes. LibTIFF's public
buffer uses the full TIFF row stride even for reduced chroma; the generator
removes that buffer padding when recording the JPEG samples.

`generate.py` uses Pillow 11.3 bilinear resizing to independently reconstruct
centered chroma, crops outside-image padding and converts YCbCr to RGB. The
`.tif.rgb` references protect every output pixel within a 3/255 tolerance for
JPEG, interpolation and conversion rounding. These are synthetic producer and
reference-rendering cases, not native Windows acceptance.

Compile `generate.c` with local LibTIFF headers/library, then run
`python3 generate.py <path-to-compiled-generator>`. LibTIFF and Pillow are
fixture tooling only; the product adds no runtime dependencies. `SHA256SUMS`
identifies all TIFF, plane and RGB reference files.

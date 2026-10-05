# Cosited JPEG TIFF fixtures

LibTIFF 4.7.2 encodes these 90 synthetic TIFF images from raw YCbCr samples.
Eighty-eight use cosited positioning; two centered chunky cases protect the
same partial-tile boundary. The separate and chunky layouts cover both byte orders, strips and tiles,
local/shared tables, odd/even image edges, and 1/1, 2/1, 2/2, 4/1 and 4/2 chroma
sampling. Separate planes additionally cover 4/4; that layout exceeds the
baseline interleaved MCU block limit when stored chunky.

`generate.c` extracts the encoded TIFF segments and uses libjpeg-turbo's raw
component API to independently decode their sample planes. This avoids a
LibTIFF raw-chroma buffer-size limitation in some tiled cases. Each `.tif.planes`
record stores five little-endian 32-bit words (plane, x, y, width, height), then
the decoded sample bytes. LibTIFF/libjpeg-turbo remain isolated fixture tools.

`generate.py` uses Pillow 11.3 affine bilinear sampling to reconstruct cosited
chroma. It crops decoded chroma to the visible image extent before extending the last
sample at boundaries, uses the TIFF sample
origin, and converts YCbCr to RGB. `.tif.rgb` files provide independent pixel
references with a 3/255 tolerance for decoder/interpolation/color rounding.
Tile padding uses contrasting chroma values to expose accidental interpolation
outside the image. JPEG compression can still affect retained samples; reference
comparisons use independently decoded samples rather than original encoder input.
Local tables use 67-by-35 images; shared tables use 68-by-36 images.

Compile the C producer with LibTIFF and libjpeg-turbo headers/libraries, then run
`python3 generate.py <compiled-generator>`. `SHA256SUMS` covers all TIFF, raw
plane and RGB reference files. These fixtures establish managed decoding and
reference sampling, not native Windows acceptance.

Sample positioning follows [TIFF 6.0 section 21](https://www.itu.int/itudoc/itu-t/com16/tiff-fx/docs/tiff6.pdf).

# JPEG TIFF extra-sample fixtures

LibTIFF 4.7.2 supplies TIFF containers and libjpeg-turbo 3.2.0 encodes the JPEG
components for these 288 35-by-19 images. The matrix covers gray polarities, RGB,
CMYK and YCbCr, both byte orders, strips/tiles, shared/local tables, chunky/separate
planes, and unspecified, associated-alpha and unassociated-alpha extra samples.
YCbCr covers 1/1 and 2/2 subsampling; the extra channel retains luma resolution.

Compile `generate.c` with LibTIFF and libjpeg-turbo headers/libraries, then run
`python3 generate.py <compiled-generator>`. The tools remain isolated fixture
tooling and are not OfficeIMO runtime requirements.

Two producer/decoder boundaries require explicit handling:

- LibTIFF rejects tiled YCbCr with an extra sample. The generator writes the raw
  JPEG data with RGB container bookkeeping, then changes only PhotometricInterpretation
  to YCbCr. The YCbCr tags and encoded component samples already describe that space.
- Libjpeg-turbo 3.2.0 encodes five-component sequential JPEG but its SOS reader
  cannot select the fifth frame component. Each unchanged single-component entropy
  scan is decoded with a one-component frame header for the reference. Whole-file
  five-component native decoder acceptance is not established by this comparison.

`.tif.raw` retains independently decoded device samples. Separate-plane `.planes`
records store five little-endian words (plane, x, y, width, height) followed by
sample bytes. Pillow reconstructs reduced centered chroma after cropping image
padding and converts YCbCr to RGB. The Python generator then applies the TIFF
extra-sample contract to produce `.rgba` references. Tests compare every alpha
sample and use RGB tolerances of 3/255, or 6/255 for associated alpha because
unassociation amplifies decoder/color rounding. Encoded alpha spans 96–239;
these fixtures do not independently qualify near-zero lossy alpha precision.

`SHA256SUMS` covers the TIFFs and reference bytes. These fixtures establish managed
component handling and transparent document rendering, not native Windows acceptance.

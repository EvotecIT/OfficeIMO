# Lossless arithmetic JPEG TIFF color and alpha

The corpus contains 168 compression-7 TIFFs with SOF11 JPEG segments at eight,
twelve and sixteen bits. It covers WhiteIsZero/BlackIsZero gray, RGB, CMYK and
full-resolution YCbCr, opaque samples, associated and straight alpha, both byte
orders, strips and tiles, and contiguous and separate planes. The matrix combines
little-endian contiguous strips, big-endian contiguous tiles, big-endian separate
strips and little-endian separate tiles. It is not every layout combination.

All seven predictors occur in the corpus. Point transforms are zero or two;
restart intervals are three complete segment rows. The 35 × 19 images exercise
partial strips/tiles and alpha values near zero and full coverage. CMYK alpha is
qualified only with separate planes: the native producer rejects five-component
frames. Subsampled YCbCr, shared JPEG tables and unspecified extra channels are
not part of this lossless-arithmetic color corpus.

## Reproduction

Build the test-only JPEG oracle at `thorfdbg/libjpeg` commit
`c719010a26ce0c666e98b2acf924ad5fc24b4f5d` after running
`../JpegArithmeticLossless/prepare_oracle.py`. The setup exposes predictor, point
transform and input component count in its driver and corrects the documented
initial predictor for point transforms. It does not change arithmetic entropy
coding. The component-count override treats PNM input as raw interleaved samples;
it does not claim standard PNM represents CMYK or alpha. The GPLv3 oracle remains
outside product builds, shipped dependencies and this repository.

Compile `wrap.c` against test-only LibTIFF 4.7.2, then run:

```sh
python3 generate.py /path/to/prepared/jpeg /path/to/wrap /task/scratch
LCMS_LIBRARY=/path/to/liblcms2 python3 ../TiffJpegArithmeticAlpha/generate-icc.py
```

The generator independently encodes and decodes every one of the 1,593 JPEG
segments and checks exact sample agreement before LibTIFF writes the raw segments.
LibTIFF bookkeeping uses RGB for YCbCr resources with extra samples; the wrapper
then restores the TIFF photometric tag. This is native segment qualification with
constructed TIFF layout/color declarations, not independent full-file decoder
acceptance. `.raw` files contain little-endian 16-bit component words at the stated
precision. `.rgba` references apply the declared TIFF color/alpha equations,
retaining fractional YCbCr-derived RGB through native-alpha unassociation and
rounding only at final byte output.
The 24 `.icc-rgba` references use LittleCMS 2.19, the existing explicit CMYK test
profile and relative colorimetric intent. Associated colorants are divided by
source-precision alpha before ICC conversion. Hashes are in `SHA256SUMS`.

Managed alpha matches exactly. Visible compositing over black and white agrees
with the independent references within 3/255, including profiled CMYK. The system
LibTIFF decoder rejects all 168 files: its JPEG build omits lossless arithmetic,
rejects sixteen-bit precision, or rejects the extra-channel YCbCr layout. Those
native reader limits remain separate from the managed decoding tests. Default
CMYK/SWOP interpretation and native Windows acceptance remain unqualified.

The 672 XPS/OpenXPS exports cover 446,880 pixel-center probes per route across
black and white backgrounds. MuPDF 1.28.2 renders normalized PDF/SVG output within
5/255 and 2/255 respectively, without warnings. GhostXPS 10.08.0 opens 656 exports
but can render blank images (255/255 error); it crashes on 16 twelve-bit CMYK strip
exports. These results do not establish general native XPS image acceptance.

# Additional-precision arithmetic TIFF alpha

These 912 TIFFs cover lossless arithmetic JPEG at 2–7, 9–11 and 13–15 bits.
The 432 full-resolution cases include WhiteIsZero/BlackIsZero gray, RGB, CMYK
and YCbCr. Another 480 YCbCr cases cover 2×1/2×2/4×2 subsampling and centered or
cosited positioning. Both sets include associated and unassociated alpha,
35×19 images, strips/partial tiles, contiguous/separate planes, and both byte
orders in four paired layouts. Five-component CMYK-alpha and four-component
4×2 YCbCr-alpha use separate TIFF planes to respect the native producer's frame
and interleaved-scan limits.

The test-only producer is `thorfdbg/libjpeg` at commit
`c719010a26ce0c666e98b2acf924ad5fc24b4f5d`, prepared by
`../JpegArithmeticLossless/prepare_oracle.py`. Its GPLv3 source and executable
remain outside the repository and shipped dependencies. The preparation exposes
existing predictor, point-transform and component-count options in the driver;
it also corrects the documented initial predictor. It does not alter arithmetic
entropy coding. Raw multi-component PNM-shaped input is a harness convention,
not a claim that PNM supports alpha or CMYK.

All seven predictors occur. Point transforms are zero or two, capped at one for
two-bit samples. Restarts align with segment rows. Full-resolution segments are
independently encoded and decoded with exact native sample agreement. For
subsampled input, the producer's decoder has an odd-height edge defect, so
matching native Huffman streams are decoded by libjpeg-turbo 3.2.0 to verify the
component planes. This is companion-stream evidence, not native arithmetic
self-decoding acceptance for those subsampled streams.

Pillow 11.3.0 provides fractional chroma interpolation. References preserve
fractional color through native-alpha unassociation before final projection.
LittleCMS 2.19 supplies 720 profile references: the existing DCI-P3 matrix profile
for RGB/YCbCr and the existing explicit CMYK LUT profile, with relative
colorimetric intent. The manifest records precision, layout, alpha kind,
predictor and point transform; SHA256SUMS covers every TIFF/reference file.
Managed comparisons cover 606,480 device-color pixels within 1/255 and 478,800
profiled pixels within 2/255, with exact alpha. XPS tests reopen both dialects
and compare black/white compositing through raster, SVG and PDF output.

Build the prepared producer, `../TiffJpegArithmeticLosslessColor/wrap.c` against
LibTIFF 4.7.2, and `../TiffJpegChroma16/decode.c` against libjpeg-turbo, then run:

```text
python3 generate.py <prepared-jpeg> <wrap> <decode> <scratch-directory>
```

Set `LCMS_LIBRARY` if needed. The shared generators accept optional precision
lists, output folders and an `alpha` selection; their default routes regenerate
the earlier 8/12/16-bit corpora. LibTIFF writes unchanged native JPEG segments;
YCbCr uses RGB bookkeeping before restoring the photometric declaration.
These are native-produced segments in constructed TIFF color/layout declarations,
not independent full-file TIFF decoder or native Windows acceptance. Wider
profiles, 4×4 arithmetic sampling and other producer/consumer combinations remain
separate qualification gaps. No generation tool is a runtime dependency.

# TIFF extra-channel fixtures

These 288 fixtures qualify one declared alpha channel among unspecified extra
samples, or no alpha channel. The three extra samples place associated alpha first
or last, unassociated alpha in the middle, or unspecified data in every position.
Multiple declared alpha channels are rejected rather than combined or selected
implicitly. TIFF 6.0 defines mixed associated and unassociated alpha as undefined.

The 192 lossless fixtures use LibTIFF 4.7.2 to write and independently decode
unsigned 8/16-bit and floating 16/24/32/64-bit samples. They cover gray polarities,
RGB, CMYK, both byte orders, chunky/separate storage, strips/tiles, and
uncompressed/LZW/Deflate/PackBits data. LZW/Deflate use the applicable integer or
floating predictor. Floating unspecified channels contain NaN, testing that only
color and declared alpha participate in finite-sample validation.

The 96 JPEG fixtures use the existing component encoder in
`../TiffJpegAlpha/generate.c`, compiled with LibTIFF and libjpeg-turbo 3.2.0.
Its optional final argument gives the extra-sample kinds, such as `020`.
They cover four-to-seven-component frames, gray/RGB/CMYK/YCbCr, both byte orders,
strips/tiles, local/shared tables, and chunky/separate storage. YCbCr includes
centered 1/1 and 2/2 chroma subsampling; all extra channels retain luma resolution.
The [existing JPEG provenance limits](../TiffJpegAlpha/README.md) apply: tiled
YCbCr metadata is finalized after raw JPEG writing, and libjpeg-turbo references
for frames with more than four components decode unchanged single-component
entropy scans through individual frame headers. Reduced chroma scans use their
actual sample dimensions before Pillow reconstruction. This is independent sample
decoding, not whole-file native acceptance of those multicomponent JPEG frames.

Compile `generate.c`, the existing JPEG generator, and
`../TiffFloating/decode.c` with their existing libraries, then run:

```text
python3 generate.py <lossless-generator> <jpeg-generator> <libtiff-decoder>
```

The Python generator converts independently decoded samples into `.rgba`
references. `.raw` and `.planes` retain device/plane samples; `SHA256SUMS` covers
all fixture and reference bytes. Tests compare all 125,856 pixels. Lossless RGB
uses a 1/255 rounding tolerance; JPEG uses 3/255, or 6/255 for associated alpha.
Alpha comparisons allow 1/255. These bounds describe this corpus, not lossless
recovery of original JPEG source pixels. Native Windows acceptance remains open.

Separate managed contract tests exercise the JPEG 255-component frame boundary,
raw-transform requirement, ordinary RGBA rejection and decoded working-set budget.
These synthetic cases do not establish independent producer acceptance at 255
components. The JPEG ten-block limit applies to each interleaved scan, not the
sum of components stored in separate scans, as specified in T.81 section B.2.3.

Sources: [TIFF 6.0 ExtraSamples](https://www.itu.int/itudoc/itu-t/com16/tiff-fx/docs/tiff6.pdf),
[T.81 JPEG](https://www.w3.org/Graphics/JPEG/itu-t81.pdf).
No generator or comparison library is an OfficeIMO runtime dependency.

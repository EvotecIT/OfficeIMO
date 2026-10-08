# Fractional TIFF JPEG chroma references

These 900 constructed TIFFs cover Huffman lossless JPEG at every precision from
2 through 16 bits. They exercise 2×1, 2×2, 4×2 and 4×4 subsampling, centered and
cosited positioning, chunky and separate planes, strips and partial tiles, and
both byte orders. Images are 5×3 or 17×11. The 4×4 cases use separate scans or
planes to respect JPEG's interleaved block limit.

The shared `../TiffJpegChroma16/generate.py` authors predictor-1 SOF3 entropy.
Libjpeg-turbo 3.2.0 independently decodes every native sample before the generator
writes references. Pillow 11.3.0 interpolates floating-point chroma on the TIFF
sampling grid; RGB conversion preserves the fractional samples until final
8-bit projection. Explicit ReferenceBlackWhite values retain native precision.
TIFF segment boundaries and partial-tile visible edges clamp their own grids.

The core test compares all reference pixels within 1/255, with exact opaque
alpha. A standalone JPEG regression extracts the centered 2×1 first strip at
each precision and checks `HighQualityChroma` color output against the same
references. The raw-component sampler retains integer rounding. XPS tests cover both package dialects and raster, SVG and PDF routes.
These fixtures reproduce visible low-precision color shifts caused by rounding
interpolated chroma before color conversion. They do not qualify subsampled
alpha, ICC profiles, arithmetic coding at the additional precisions, or native
Windows acceptance. They are specification-authored inputs with independent
sample decoding and interpolation, not independently produced TIFF files.

Compile `../TiffJpegChroma16/decode.c` against libjpeg-turbo, then run:

```text
python3 generate.py <compiled-decoder> <scratch-directory>
```

`manifest.csv` records dimensions, sampling and layouts. `.rgba` files contain
reference pixels; `SHA256SUMS` identifies TIFF and reference bytes. Generation
uses test-only tools and adds no runtime dependency to OfficeIMO.

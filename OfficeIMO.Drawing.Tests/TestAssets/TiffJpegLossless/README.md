# Huffman lossless JPEG in TIFF

These 224 TIFF fixtures contain eight-bit Huffman lossless JPEG (SOF3) encoded and
independently decoded by libjpeg-turbo 3.2.0. LibTIFF 4.7.2 writes the TIFF container
through its raw write API. A second full-file decode through LibTIFF's
TIFFReadScanline/TIFFReadTile APIs agrees exactly on all 224 files.

The matrix covers all seven predictors; point transforms 0, 1, 3 and 7;
interleaved or separate component scans; no restart or two-row restart intervals;
both TIFF byte orders; strips/tiles; chunky/separate planes; both grayscale
polarities, RGB and CMYK. Gray and RGB include unassociated alpha. Each page is
35 by 19 pixels; edge tiles retain their padded source dimensions.

`generate.c` uses the libraries' public APIs. Its decoder verifies every component
against the original source with low point-transform bits cleared. `.tif.raw`
contains the independently decoded page components, and `.rgba` is the Python
projection into the existing TIFF color/alpha contract. The first encoded segment
and its independently decoded components are retained as `.tif.jpg` and
`.tif.jpg.raw`, covering the shared JPEG owner independently of TIFF assembly.
`SHA256SUMS` covers the fixture and reference bytes.

Compile `generate.c` with the existing LibTIFF and libjpeg-turbo headers/libraries,
and `../TiffFloating/decode.c` with LibTIFF, then run:

```text
python3 generate.py <compiled-generator> <compiled-libtiff-decoder>
```

The public TIFF path compares every RGBA sample exactly across 148,960 pixels.
The direct JPEG path compares every component exactly. Additional managed tests
reject invalid predictors, scan fields, quantization selectors, AC selectors,
missing tables, truncated entropy, misaligned restarts and wrong restart sequence;
they also exercise cancellation and retained-memory bounds. Higher precision,
arithmetic coding and native Windows acceptance remain unqualified.

The decoder follows [T.81 Annex H](https://www.w3.org/Graphics/JPEG/itu-t81.pdf).
Point transforms discard low bits: only point transform zero is lossless with
respect to all original source bits. Generator and comparison libraries are
validation tools and are not OfficeIMO runtime dependencies.

Sixteen additional `edges/` JPEG fixtures are authored from T.81 sample/entropy
rules and independently decoded by libjpeg-turbo `djpeg`. They cover constant
128 samples with horizontal/vertical subsampling, even/odd dimensions, and
separate/interleaved scans. Twelve matching TIFF wrappers use supported TIFF
subsampling; vertical-only JPEG subsampling has no TIFF wrapper. These are
independent-decoder checks, not independent-producer evidence. Run
`python3 generate-edges.py <djpeg>` to regenerate them. They prevent interpolation
from reading zero padding past the right/bottom lossless component edge.

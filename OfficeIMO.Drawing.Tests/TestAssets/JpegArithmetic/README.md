# Sequential arithmetic JPEG fixtures

Ninety JPEGs are independently encoded and decoded with libjpeg-turbo 3.2.0.
They cover 35×19 gray, RGB and YCbCr images; eight/twelve-bit samples;
quality 1/75/100; 1×1, 2×1 and 2×2 chroma; partial blocks; and interleaved
or separate scans. Restart cases use two MCUs per interval with separate scans or three with
interleaved scans, DC destination 15,
AC destination 14, and non-default conditioning L=2, U=5, K=12. Other cases use
default conditioning. Native integer slow IDCT references include nearest and
high-quality chroma reconstruction, projected to RGBA8 after native decoding.

LibTIFF 4.7.2 also wraps each encoded stream in a compression-7, little-endian,
single-strip TIFF. This is independent container production with a native JPEG
pixel reference, not independent full-file TIFF decoder acceptance. TIFF comparisons
use the high-quality chroma reference. CMYK, alpha, planar/tiled arithmetic TIFF,
and progressive/lossless arithmetic JPEG have no qualification in this corpus.

Build `generate.c` against libjpeg-turbo and LibTIFF headers and link `-ljpeg -ltiff`.
Run `python3 generate.py /absolute/path/to/generate`. These tools are test-only;
normal builds and tests consume the retained fixtures without either dependency.
`manifest.csv` records cases; `SHA256SUMS` covers images and pixel references.

The managed arithmetic implementation follows [ITU-T T.81](https://www.w3.org/Graphics/JPEG/itu-t81.pdf),
Annexes D and F, including the probability state table, conditioning and restart reset.

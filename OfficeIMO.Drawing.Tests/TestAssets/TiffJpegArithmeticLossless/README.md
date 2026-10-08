# Arithmetic lossless JPEG in TIFF

These 168 compression-7 TIFF files contain eight/twelve/sixteen-bit grayscale or
RGB SOF11 payloads. LibTIFF 4.7.2 writes both byte orders with a single chunky
strip. The JPEG payloads come from the [lossless arithmetic corpus](../JpegArithmeticLossless/README.md),
which documents the native producer, its explicit corrections, and independent
sample calibration. This container corpus covers all seven predictors, point
transforms 0 and precision-minus-one, and row-aligned restarts.

Core validates every container and reproduces all 35,112 expected RGBA pixels
exactly. TIFF owns the photometric interpretation and byte order; JPEG application
markers do not override those tags. The shared decoder preserves twelve/sixteen-bit
samples until TIFF color conversion and eight-bit output. Legacy compression 6
retains its separate Huffman process contract.

## Regeneration

`wrap.c` uses LibTIFF only to write the container and copy the existing JPEG bytes
through `TIFFWriteRawStrip`. It does not encode or decode JPEG, and it is not a
product or ordinary-test dependency. Build it with an existing test-only LibTIFF
installation, then generate and verify:

```sh
cc wrap.c $(pkg-config --cflags --libs libtiff-4) -o /task/scratch/wrap
python3 generate.py /task/scratch/wrap
shasum -a 256 -c SHA256SUMS
```

The source JPEG files must already exist in the adjacent corpus. `manifest.csv`
records each container's source file, precision, predictor, point transform,
restart interval and byte order.

## Evidence boundaries

LibTIFF independently writes these containers, but its JPEG decoder does not
accept them in this environment: 112 eight/twelve-bit cases report an omitted
codec feature, and 56 sixteen-bit cases report improper JPEG precision. This is
independent container-production and native JPEG-sample evidence, not native
full-file TIFF decoder acceptance.

The 336 XPS/OpenXPS exports cover 70,224 pixel-center probes per route. MuPDF 1.28.2
renders PDF/SVG within 2/255 of managed output without warnings. GhostXPS 10.08.0
opens every package but differs by up to 255/255, including blank images.
Arithmetic TIFF CMYK/YCbCr/alpha, planar/tiled layouts, and native Windows rendering
remain unqualified by this corpus.

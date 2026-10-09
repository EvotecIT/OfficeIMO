# Twelve-bit lossless JPEG in TIFF

These 280 fixtures use libjpeg-turbo 3.2.0 to encode and independently decode
native twelve-bit SOF3 samples. LibTIFF 4.7.2 writes their containers through its
raw segment API. The shared generator is in `../TiffJpegLossless16`: compile
`generate.c` against libjpeg-turbo and LibTIFF, then run
`python3 ../TiffJpegLossless16/generate.py /absolute/path/to/generate 12 .`.
These native libraries are opt-in verification tools, never runtime dependencies.

The 35×19 matrix covers all seven predictors, point transforms 0/1/6/11,
separate/interleaved scans, row-aligned restarts, both byte orders, chunky/planar
strips/tiles, both gray polarities, RGB, CMYK and full-resolution YCbCr. Gray,
RGB and YCbCr include associated/straight alpha and native alpha values 0 and 1.

`.tif.jpg` retains the first compressed segment; `.jpg.raw` and `.tif.raw` contain
independently decoded little-endian sample words. The producer verifies every word
against the original samples with discarded point-transform bits cleared.
`.rgba` projects the decoded words through declared color and alpha equations.
`.reference.tif` rescales non-YCbCr words to unsigned sixteen-bit storage for
ICC comparison; that rescaling can introduce one output level of rounding.
YCbCr references use normalized floating-point RGB to preserve fractional color
through conversion and unassociation. The original compressed words are unchanged.
`SHA256SUMS` records all fixture and reference bytes.

`decode-native.c` checks full-file LibTIFF consumption independently of the
segment oracle. It accepts 252 files: 182 match exactly and 70 differ at their
last pixel because its twelve-bit packing loop omits an odd final sample.
The other 28 files are chunky YCbCr with an alpha channel; LibTIFF rejects their
strip/tile sizing before sample decoding. These are explicit full-file native
qualification gaps. The raw JPEG decoder and managed TIFF paths cover all 280.

YCbCr is written through an RGB-shaped raw container and its PhotometricInterpretation
tag is changed to YCbCr; this is declared container construction, not an unchanged
native YCbCr-alpha producer. Component words and compressed JPEG packets are not
modified. Wider subsampling, arithmetic JPEG, legacy twelve-bit table-pointer
layouts and native Windows acceptance remain outside this corpus.

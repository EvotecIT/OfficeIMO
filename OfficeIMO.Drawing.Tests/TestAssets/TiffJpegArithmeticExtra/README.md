# Unspecified extra samples in lossless arithmetic TIFF

These 114 TIFFs verify that one extra channel declared `ExtraSamples=0` is ignored:
colors remain unchanged and every output pixel is opaque. They cover eight,
twelve and sixteen bits; gray, RGB, CMYK and YCbCr; full-resolution and subsampled
chroma; strips/tiles, both byte orders, and contiguous/separate sample layouts.
The six CMYK cases use separate planes and an explicit ICC profile. Coverage
inherits the layout limits of the source corpora; it does not add chunky CMYK or
multi-scan chunky 4×2 YCbCr extras.

`generate.py` derives the files from the independently referenced straight-alpha
cases in `TiffJpegArithmeticLosslessColor` and `TiffJpegArithmeticLosslessChroma`.
It verifies source hashes, changes only the inline extra-sample declaration from
2 to 0, and asserts that exactly one byte differs. JPEG streams remain untouched.
Reference RGB bytes remain unchanged, including pixels where the ignored channel
is zero; only reference opacity becomes 255. The manifest records source files
and hashes. Six LittleCMS reference projections are inherited unchanged in color.

```sh
python3 generate.py
```

These are deliberately modified declarations over qualified sample streams, not
an independent producer corpus for unspecified extras. Tests require exact opacity
and color agreement within 3/255, including explicit-profile CMYK. The system
LibTIFF reader fails on all 114 files: 56 missing-codec, 28 precision and 30
extra-channel layout failures. Independent full-file acceptance and native Windows
acceptance remain separate gaps. `SHA256SUMS` identifies the retained bytes.

The 456 XPS/OpenXPS exports cover 303,240 pixel-center probes per route. MuPDF
1.28.2 renders normalized PDF/SVG within 2/255 without warnings. GhostXPS 10.08.0
opens 452 packages but can render blank images (255/255 error), and crashes on
four twelve-bit CMYK strip exports. These are consumer limits, not native
acceptance of the full corpus.

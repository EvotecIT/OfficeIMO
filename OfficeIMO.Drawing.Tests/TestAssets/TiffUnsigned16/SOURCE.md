# Unsigned sixteen-bit TIFF corpus

These 47 synthetic TIFF images are independently encoded by LibTIFF 4.7.2. The test-only C producer reads every strip or tile back through LibTIFF and compares the unsigned sixteen-bit samples with the originals before writing the manifest and reference RGBA files. OfficeIMO is not used to produce the fixtures or their reference pixels.

The corpus covers little- and big-endian classic TIFF, uncompressed/LZW/PackBits/Deflate payloads, horizontal prediction, chunky/planar strips and padded tiles, gray/RGB/CMYK samples, associated/unassociated/unspecified extra samples, and a two-page file. Low associated alpha includes sample values 1, 129 and 257 to detect conversion that rounds alpha before unpremultiplication. Palette images with sixteen-bit indices, mixed component widths, reversed FillOrder, signed/floating/undefined samples, JPEG compression and BigTIFF are outside this decoder contract.

Embedded RGB and CMYK profiles reuse the licensed test profiles in [IccColorCorpus](../IccColorCorpus/SOURCE.md). The gray profile is the existing synthetic gamma-1.8 profile in `OfficeIMO.Xps.Tests/Fixtures/ColorImages`. LittleCMS 2.19 produces the ICC reference pixels with a version-2.1 sRGB destination, relative-colorimetric intent, no black-point compensation, and optimization/cache disabled. ICC pixel tests allow two channel values for interpolation and rounding. Unprofiled RGBA samples must match exactly. Unprofiled CMYK references describe the ordinary Core decoder's existing device approximation; XPS requires a usable ICC profile for those images.

LibTIFF and LittleCMS are optional fixture-generation tools. They are not shipped or required by OfficeIMO or its normal tests. Run from the repository root:

```sh
cc OfficeIMO.Drawing.Tests/TestAssets/TiffUnsigned16/generate_libtiff.c \
  $(pkg-config --cflags --libs libtiff-4 lcms2) -lm -o /tmp/tiff16-producer
/tmp/tiff16-producer OfficeIMO.Drawing.Tests/TestAssets/TiffUnsigned16 \
  OfficeIMO.Drawing.Tests/TestAssets/IccColorCorpus/littlecms-rgb-matrix.icc \
  OfficeIMO.Drawing.Tests/TestAssets/IccColorCorpus/littlecms-cmyk-lut.icc \
  OfficeIMO.Xps.Tests/Fixtures/ColorImages/gray-gamma18.icc
```

`manifest.csv` identifies the image model and page count. The `.rgba` files contain nineteen by thirteen straight-alpha eight-bit reference pixels, after ICC conversion when a profile is embedded. Quantization occurs only at the output boundary. Associated sample components are divided by the full sixteen-bit alpha, with zero-alpha components set to zero. The producer source explains the deterministic sample formulas. `SHA256SUMS` records the checked-in binary corpus.

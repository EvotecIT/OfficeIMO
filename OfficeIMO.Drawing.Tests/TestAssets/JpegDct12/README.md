# Twelve-bit DCT JPEG references

The 60 fixtures cover twelve-bit extended sequential (SOF1) and progressive (SOF2)
JPEG with gray, RGB and YCbCr pixels. Libjpeg-turbo 3.2.0 independently encodes and
decodes every file. The matrix includes quality 1/75/100, 1×1/2×1/2×2 YCbCr sampling,
zero/two-row restart intervals, interleaved/separate sequential scans, progressive
refinement and 35×19 dimensions that exercise partial blocks. Quality 1 permits
wide quantization values. Input ramps, high-frequency changes and black/white
edges exercise transform range and saturation.

The decoder uses native twelve-bit output with `JDCT_ISLOW`, with fancy chroma
upsampling disabled and enabled. `.nearest.rgba` and `.bilinear.rgba` project its
native RGB/gray samples to eight-bit RGBA with nearest rounding. Managed output
must agree within one channel value across 79,800 pixel comparisons in the two quality modes. The public output is eight-bit; these
references qualify its pixels, not a public twelve-bit sample API.

Compile `generate.c` against an existing libjpeg-turbo installation, then run:

```text
python3 generate.py <compiled-helper>
```

`manifest.csv` describes each file. `SHA256SUMS` identifies the fixture/reference
bytes. These are opt-in test tools; normal OfficeIMO builds and runtime packages
have no native JPEG dependency. Baseline SOF0 remains eight-bit. Arithmetic JPEG,
packed twelve-bit TIFF samples and native Windows XPS acceptance are separate work.

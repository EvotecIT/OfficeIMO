# Subsampled TIFF chroma, alpha and profile references

These 1,680 specification-authored TIFFs cover every precision from 2 through
16 bits, associated and unassociated alpha, 2×1/2×2/4×2/4×4 chroma grids, centered
and cosited positioning, chunky and separate planes, strips and partial tiles,
and both byte orders. Luma and alpha stay at full resolution. Images are 5×3 or
17×11; layouts that would exceed JPEG's ten-sample interleaved MCU limit use
separate scans or planes instead.

The shared `../TiffJpegChroma16/generate.py` authors predictor-1 SOF3 streams.
Libjpeg-turbo 3.2.0 independently verifies all native words in 6,840 JPEG segments.
Pillow 11.3.0 performs fractional chroma interpolation with visible-edge clamping.
The reference converts YCbCr to RGB and unassociates with native alpha before
clipping or quantization. The alpha pattern includes zero, one, partial and full
coverage. Associated luma uses per-pixel alpha; associated chroma averages
alpha over its sampling group before native quantization. Cases include nonzero native alpha that rounds to zero in eight-bit output.

`generate.py` sends normalized floating-point RGB to LittleCMS 2.19 for explicit
DCI-P3 matrix-profile conversion to sRGB, using relative colorimetric intent.
The profile changes channels by up to 133/255. Device and profiled comparisons
each cover 169,680 pixels: alpha matches exactly, device RGB within 1/255 and
profiled RGB within 2/255. XPS tests reopen both package dialects and compare
transparent alpha and black/white compositing through raster, SVG and PDF routes.

Compile `../TiffJpegChroma16/decode.c` against libjpeg-turbo, then run:

```text
python3 generate.py <compiled-decoder> <scratch-directory>
```

Set `LCMS_LIBRARY` if LittleCMS cannot be discovered. Tools and intermediate
floating-point files are used only for opt-in fixture generation, not runtime
or normal builds. `manifest.csv` records layout and color parameters;
`icc-reference.json` identifies the profile, intent and LittleCMS version.
`SHA256SUMS` covers the TIFFs and both RGBA reference sets.

These constructed files provide independently decoded sample and color evidence,
not independent TIFF producer or native Windows acceptance. Native multi-scan producer evidence, further profiles and wider native consumers
remain separate qualification gaps.

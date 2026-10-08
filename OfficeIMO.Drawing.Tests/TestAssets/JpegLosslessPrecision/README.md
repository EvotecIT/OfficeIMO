# Lossless JPEG sample precision

This corpus contains 155 lossless JPEG images with 17×11 pixels. The opt-in
libjpeg-turbo 3.2.0 producer/decoder checks every native sample before writing a
reference. No native library is required by OfficeIMO or its ordinary tests.

The 140 gray/RGB cases cover every precision from 2 through 16 bits, predictors
1 and 7, zero and maximum point transforms, two-row restart intervals and
interleaved/separate scans. Twelve-bit cases cover all seven predictors.
The 15 YCbCr cases use independently encoded raw components with conventional
component IDs 1/2/3. They include neutral chroma and varying colors at every
precision. Their native samples are independently decoded; RGB references use
the YCbCr equations, not a claim of native RGB decoder agreement. Chroma is
centered at 2^(precision-1) before projection and clipping. This matters especially
at low precision, where projecting each component first shifts the neutral value.

`.raw` files contain little-endian native component words. `.rgba` files contain
the equation-derived YCbCr reference. Gray/RGB projection maps zero to zero and
the inclusive precision maximum to 255, with nearest rounding. A nonzero point
transform discards low bits; reference equality does not recover those bits.
The corpus checks 60,775 native samples. All new fixtures use full-resolution
components; both nearest and high-quality rendering modes are checked. Invalid precisions,
cancellation and retained-memory limits are tested separately.

Compile `generate.c` against an existing libjpeg-turbo installation, then run:

```text
python3 generate.py <compiled-helper>
```

`manifest.csv` records each case; `SHA256SUMS` identifies all image/reference bytes.
The producer uses raw-component mode for the YCbCr cases because the native
lossless colorspace path does not provide a reliable round trip for this input.
Thus these files are component-produced compatibility evidence, not a photographic
YCbCr producer corpus. TIFF packed sample widths, twelve-bit DCT, arithmetic JPEG
and native Windows XPS acceptance remain separate qualification work.

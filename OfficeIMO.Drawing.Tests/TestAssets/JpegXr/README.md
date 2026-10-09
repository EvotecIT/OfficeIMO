# JPEG XR decoder fixtures

These images exercise the public managed gray/RGB JPEG XR decoder. `manifest.csv`
contains the dimensions and pairs each `.jxr` with straight RGBA reference pixels.
`provenance.json` records producer settings and SHA-256 hashes.

The 70 ordinary images were encoded and decoded with the ITU T.835 (08/2016)
reference software from its published archive. Input pixels are generated patterns,
not third-party photographs. The cases cover spatial/frequency order, all overlap
modes, cropped macroblocks, wide rows, hard/soft tiles, gray/RGB, two alpha layouts,
lossy quantizers, trimmed flexbits, and reduced subbands. No reference executable
or library is required by the tests or shipped packages.

For separate alpha, the producer's AlphaByteCount field incorrectly contains the
whole file size. The fixtures normalize that one field to the bytes remaining
from AlphaOffset. The normalized files were independently decoded again; encoded
pixel packets were unchanged.

`rgba-premultiplied` encodes known associated samples. Its pixel-format GUID and
codestream premultiplication flag declare those samples as PBGRA. Its expected
buffer tests the public straight-alpha output contract, including zero alpha.
`rgba-premultiplied-separate` encodes the same associated samples in separate
primary and alpha streams, with the older clear premultiplication flags. Tests
also exercise the modern alpha flag and reject contradictory straight-alpha
declarations. Both variants are recorded separately from the independently
decoded ordinary corpus.

The test-only helper `OfficeIMO.TestAssets/JpegXrTestFixture.cs` creates orientation
and embedded-profile variants by appending a replacement directory. It retains
all original encoded packets and their offsets.

The unsigned-sixteen-bit fixtures use generated gray/RGB/RGBA TIFF samples and
independent raw sixteen-bit reference decoding. Cases cover both packet orders,
overlap modes, hard tiles, lossy quantization, separate/interleaved alpha, and
SHIFT_BITS values of 1, 7, and 15. Shift cases modify the sample-shift header and
are decoded independently again. PRGBA variants encode known associated samples;
expected RGBA8 values unassociate at sixteen-bit precision before rounding.
The corpus also protects agreement between container and codestream sample depth.
The retained sixteen-bit TIFF source images also check ICC conversion against
the existing source-precision TIFF path. They are original generated patterns.

The subsampled fixtures cover 4:2:0 and 4:2:2 chroma, both packet orders,
eight/sixteen-bit samples, overlap, hard/soft tiles, reduced bands, trimmed
flexbits, and non-default sampling-grid centering. Expected pixels come from
the ITU reference decoder. The wider comparison covers 274 images; the Microsoft
comparison decoder differs on non-default centering and by at most one alpha
level on lossy eight-bit interleaved-alpha cases. Those differences are not
treated as exact agreement between independent consumers.

The extended-sample fixtures cover signed s2.13/s7.24 fixed-point, IEEE half/single
floating-point, gray/RGB/RGBA, both packet orders, hard tiles, lossy quantization,
large and small finite values, and premultiplied float alpha. The expected RGBA
buffers apply scRGB-to-sRGB conversion to independently decoded source samples,
unassociating before conversion where declared. Selected `.linear` files retain
those normalized source values as little-endian IEEE doubles for ICC integration
tests. NaN/infinity fixtures protect rejection without partial pixel output.

The broader extended-sample comparison checks 151 files against raw ITU output
before RGBA rounding. Microsoft comparison decoding differs on lossy
interleaved-alpha samples; this is retained as a consumer qualification limit.

Signed endpoint fixtures distinguish sixteen-bit clipping from thirty-two-bit
output packing when lossy reconstruction crosses a representable endpoint.

The CMYK corpus in `cmyk-manifest.csv` carries independent decoded source samples
in `.cmyk` files (interleaved C/M/Y/K/optional alpha, unsigned eight-bit or
little-endian sixteen-bit). Tests convert these samples through the existing ICC
engine before comparing image and PDF-reader pixels. The wider comparison covers
96 encodings and 768,768 exact channel samples. Microsoft comparison decoding
accepts ordinary CMYK, with one-level alpha differences in four lossy eight-bit
interleaved-alpha cases, and rejects CMYKDirect identifiers.

The ITU encoder input adapter pads the final eight-bit interleaved-alpha TIFF
strip to a macroblock boundary; its reader requests a full strip after the image
height. CMYKDirect container GUIDs are corrected to match the encoded output color
format, and separate-alpha byte counts include the complete alpha stream. Neither
correction changes pixel packets. Corrected files are decoded again by the reference
program; its planar CMYKDirect output is interleaved for `.cmyk` storage. Fixture
hashes and producer details are recorded in `provenance.json`.

`nchannel-manifest.csv` covers three through eight unsigned 8/16-bit color
channels. The `.nchannel` files contain interleaved source channels and optional
alpha, with sixteen-bit samples stored little-endian. All 422,730 samples in 60
fixtures match the ITU reference decoder. Of these, 48 are direct encodings and
12 are containers assembled from independently encoded primary and grayscale
alpha streams because the reference encoder asserts for separate N-channel
alpha. The reference decoder reads each assembled container again. No encoded
pixel packets are changed. Tests exercise matching profiles, rejected mismatches,
resource limits, both XPS dialects, SVG pixels, and PDF-reader alpha. Microsoft
comparison decoding differs on interleaved alpha and one eight-channel frequency
case; these fixtures do not establish complete native interoperability.

The 24 `mixed-*` fixtures cover all six valid reduced interleaved-alpha band
combinations at unsigned 8/16-bit precision in spatial and frequency packet
order. Microsoft jxrlib produces them with a configuration-only adapter exposing
its alpha-plane `sbSubband` setting; encoding algorithms are unchanged. The
unmodified ITU decoder supplies the twelve spatial RGBA references. Each frequency
case uses its equivalent spatial encoding's reference, with identical input and
quantization. Native frequency decoders fail or disagree for these cases, so
paired-packet equivalence does not establish direct native frequency acceptance.
`provenance.json` records the source archive hash and each reference pairing.

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

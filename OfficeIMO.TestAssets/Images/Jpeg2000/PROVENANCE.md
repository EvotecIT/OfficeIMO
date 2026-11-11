# JPEG 2000 decoder fixtures

Source: SerenityOS `Tests/LibGfx/test-inputs/jpeg2000` at commit
`e70c5343579db8c952c6092e16610e6886dadb24`
(https://github.com/SerenityOS/serenity), BSD-2-Clause, Copyright (c) the SerenityOS developers.
The files were produced by independent encoders (JasPer, Kakadu and OpenJPEG) from one 119x101 source image.

`manifest.json` pins each file by SHA-256 and records the feature it exercises. `samplesSha256` is the SHA-256
of the eight-bit interleaved samples decoded by `opj_decompress` from OpenJPEG 2.5.4 (Homebrew), with alpha last.
Every lossless RGBA file therefore shares one expected hash. Irreversible files are compared with the lossless
decode of the same source; OpenJPEG's own decodes of them differ from it by at most 3.

The ISO/IEC 15444-4 conformance codestreams and the PDF corpus files used for wider validation are not committed;
their copyright terms limit use to conformance testing. See the JPX section of
`Docs/benchmarks/pdf-jpx-jbig2-decoding-2026-11-10.md` for those runs.

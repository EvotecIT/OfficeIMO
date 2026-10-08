# Third-party notices

## Hyphenation patterns

The embedded US English and reformed German resources are derived from
[`hyphenation/tex-hyphen` at `5684c0f51c0b81133db2efbe60a408b4155a3ff5`](https://github.com/hyphenation/tex-hyphen/tree/5684c0f51c0b81133db2efbe60a408b4155a3ff5/hyph-utf8/tex/generic/hyph-utf8/patterns/tex).
Pattern tokens and explicit exceptions are unchanged. The resource manifest records source and
derived-file hashes. Original notices are retained in the embedded data and in
`Licenses/hyphenation-en-us-LICENSE.txt` and `Licenses/hyphenation-de-1996-LICENSE.txt`.

US English patterns are copyright Gerard D.C. Kuiken and permit redistribution with the copyright
and permission notice retained. Reformed German patterns are copyright the named
Deutschsprachige Trennmustermannschaft authors and use the MIT license.
The resources introduce no TeX or third-party hyphenation runtime dependency.

## Bouncy Castle C# AES implementation

`Security/ManagedAesBlockCipher.cs` is adapted from `AesLightEngine.cs` in Bouncy Castle C# release 2.7.0:

- Source: https://github.com/bcgit/bc-csharp/blob/release-2.7.0/crypto/src/crypto/engines/AesLightEngine.cs
- Source Git blob: `bd8bb4468588fcb92d350881aec2ae04f5f76b73`
- License: https://github.com/bcgit/bc-csharp/blob/release-2.7.0/LICENSE.md

MIT License (https://opensource.org/licenses/MIT)

Copyright (c) 2000-2026 The Legion of the Bouncy Castle Inc. (https://www.bouncycastle.org).
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sub license, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions: The above copyright notice and this
permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT
OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


## WebM VP8 reference

The managed VP8 decoder follows RFC 6386 and was checked against the WebM libwebp reference. The WebM reference license and patent grant are retained in `Licenses/libvpx-LICENSE.txt` and `Licenses/libvpx-PATENTS.txt`. No libwebp or libvpx runtime dependency is introduced.

## CodeGlyphX VP8 encoder

`Raster/Webp/OfficeVp8Encoder*.cs` and `Raster/Webp/OfficeVp8BoolEncoder.cs` are adapted from the managed CodeGlyphX VP8 encoder at [`3c9103da07623e15804450ec0bca284bfcbf4a2c`](https://github.com/EvotecIT/CodeGlyphX/tree/3c9103da07623e15804450ec0bca284bfcbf4a2c/CodeGlyphX/Rendering/Webp), copyright Przemyslaw Klys. These files retain the Apache-2.0 license in [`Licenses/CodeGlyphX-LICENSE.txt`](Licenses/CodeGlyphX-LICENSE.txt). OfficeIMO's changes split codec responsibilities, reuse Core prediction and reconstruction, bound output and managed memory, observe cancellation, and correct alpha prediction at image boundaries.

The forward integer transform in `OfficeVp8Encoder.Transform.cs` is adapted from WebM libvpx and retains its BSD-3-Clause license and patent grant in [`Licenses/libvpx-LICENSE.txt`](Licenses/libvpx-LICENSE.txt) and [`Licenses/libvpx-PATENTS.txt`](Licenses/libvpx-PATENTS.txt). The previously incorporated Core VP8 decoder remains under its existing MIT license; that relicensing does not apply to the copied encoder. These source adaptations add no CodeGlyphX, libwebp, or libvpx runtime dependency.

## ImagePlayground HEIF container metadata

`Raster/Heif/OfficeHeifMetadataReader*.cs` and `Raster/Heif/OfficeHeifImageInfo.cs` are adapted
from Evotec's ImagePlayground HEIF metadata implementation, copyright (c) 2022 Evotec, under
the MIT license. The original permission and copyright notice are retained in
[`Licenses/ImagePlayground-LICENSE.txt`](Licenses/ImagePlayground-LICENSE.txt). The adaptation
adds bounded byte and stream APIs, cancellation, collection and output limits, shared-extent
protection, and framed metadata replacement. Core has no ImagePlayground runtime dependency.

## AV1 reference data

The managed AVIF decoder implements a bounded AV1 still-image subset. Generated default probability tables are checked against hash-pinned AOM v3.13.1 numeric data; quantizer and prediction tables are generated from the AV1 1.0.0 Errata 1 specification. The AOM reference license and patent notice are retained in `Licenses/aom-LICENSE.txt` and `Licenses/aom-PATENTS.txt`.

Fixture generators under `OfficeIMO.Drawing.Tests/TestAssets/Avif` use isolated native AOM, libavif and dav1d reference tools. They are opt-in test tooling and retain the downloaded references' notices beside their sources. OfficeIMO.Core adds no native codec or external runtime package dependency.

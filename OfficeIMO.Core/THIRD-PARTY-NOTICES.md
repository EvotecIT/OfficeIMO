# Third-party notices

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

## AV1 reference data

The managed AVIF decoder implements a bounded AV1 still-image subset. Generated default probability tables are checked against hash-pinned AOM v3.13.1 numeric data; quantizer and prediction tables are generated from the AV1 1.0.0 Errata 1 specification. The AOM reference license and patent notice are retained in `Licenses/aom-LICENSE.txt` and `Licenses/aom-PATENTS.txt`.

Fixture generators under `OfficeIMO.Drawing.Tests/TestAssets/Avif` use isolated native AOM, libavif and dav1d reference tools. They are opt-in test tooling and retain the downloaded references' notices beside their sources. OfficeIMO.Core adds no native codec or external runtime package dependency.

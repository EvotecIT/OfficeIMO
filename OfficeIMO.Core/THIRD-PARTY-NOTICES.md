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


## Managed VP8 keyframe decoder

`Raster/Webp/OfficeVp8*.cs` is adapted from CodeGlyphX at commit `fc25e2fcf795d9c9a09b88708c47bdfeed5c446d`. The OfficeIMO adaptation removes encoder and diagnostic scaffolding, splits decoding responsibilities, and adds cooperative cancellation and aggregate memory guards. It includes the bounded arithmetic-partition termination correction also contributed to CodeGlyphX in commit `ad84d4be80ce5395ec655a8e1b149c6b11767942`.

- Source: https://github.com/EvotecIT/CodeGlyphX/tree/fc25e2fcf795d9c9a09b88708c47bdfeed5c446d/CodeGlyphX/Rendering/Webp
- Copyright CodeGlyphX contributors
- License: Apache License 2.0; see `Licenses/CodeGlyphX-Apache-2.0.txt`.
- Entropy, prediction, reconstruction and filter behavior follows RFC 6386 and was checked against the WebM libwebp reference. The WebM reference license and patent grant are retained in `Licenses/libvpx-LICENSE.txt` and `Licenses/libvpx-PATENTS.txt`.

No CodeGlyphX, libwebp or libvpx runtime dependency is introduced.

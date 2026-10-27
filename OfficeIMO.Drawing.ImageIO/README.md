# OfficeIMO.Drawing.ImageIO

Apple system raster decoding for OfficeIMO.Drawing, including HEIC/HEIF still-image tiles. The package calls ImageIO, CoreGraphics and CoreFoundation directly and adds no third-party runtime decoder or native binary.

The `net8.0` and `net10.0` assemblies use the same framework imports on macOS and iOS. Building this library needs no Apple workload. Running it requires macOS 10.13+, iOS 11+, or Mac Catalyst 13.1+; the device and installed OS determine which HEVC profiles decode. iOS execution needs qualification in the consuming app.

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.ImageIO;

// Register once at application startup. Other operating systems are unchanged.
OfficeImageIOCodec.Register();

// Alternatively, pass an explicit codec for a particular decode request.
var options = new OfficeRasterDecodeOptions {
    ImageCodec = new OfficeImageIOCodec()
};
```

`IsSupported` and `GetAvailability()` describe the supported operating system. They do not guarantee support for every HEVC profile. `TryDecode` returns `false` and a null image for malformed, unsupported or over-budget input. The content type is advisory; ImageIO identifies the bytes.

## HEIF composition and sample encoding

Core composes the full container. It supplies the codec with a synthesized single `hvc1` item carrying its `hvcC` and stored dimensions, without `irot`, `imir`, `clap`, `colr`, EXIF or auxiliary-image properties. Core handles grids, crop, orientation, alpha auxiliaries and ICC/nclx conversion. Call the shared Core decoder for complete HEIF files; calling this tile codec directly on an arbitrary original HEIF file does not provide that composition contract.

The codec returns tightly packed, top-to-bottom, 8-bit **straight-alpha RGBA in the decoded source RGB encoding**, before display gamut/transfer matching. It does not return raw YCbCr, linear RGB, or necessarily sRGB. Native HEVC YCbCr reconstruction, chroma upsampling, range/matrix interpretation and quantization to 8 bits still occur inside Apple's decoder. Higher-bit-depth samples are reduced to 8 bits. HDR gain-map composition is outside this codec.

The extraction path is deliberate:

1. `CGImageSourceCreateImageAtIndex` reads the stored image, with floating-point output and decoded caching disabled.
2. RGB images are retagged using `CGImageCreateCopyWithColorSpace` and `CGColorSpaceCreateDeviceRGB`. The replacement changes the image's color-space label before drawing.
3. The RGB image is drawn into a bitmap using that same DeviceRGB space, 8 bits/component, four bytes/pixel, `kCGBitmapByteOrder32Big | kCGImageAlphaPremultipliedLast` (`0x4001`), copy blending and no interpolation. No explicit geometric transform is applied.
4. Native monochrome images use the equivalent DeviceGray path: 8-bit gray plus premultiplied alpha, two bytes/pixel, `kCGImageAlphaPremultipliedLast` (`0x1`), followed by replication of the gray sample into R/G/B. This supports monochrome alpha tiles without a gray-profile conversion.
5. Premultiplied RGB is converted to straight alpha; fully transparent pixels have zero RGB. Opaque tile samples avoid this rounding step.

This relies on ImageIO exposing decoded RGB samples in the CGImage's declared source space and on color-space replacement preserving those component samples. Apple's [color-space copy API](https://developer.apple.com/documentation/coregraphics/cgimage/copy(colorspace:)?language=objc) documents replacement of the color space; it does not guarantee the entire HEVC decoder's internal color pipeline. Native tests use known sRGB and Display-P3 PNG samples to check exact sample preservation and distinguish it from an sRGB-matched reference. HEIC tests compare original system-encoded controls with a separate color-managed native full-file decoder. These are empirical controls on the tested OS, not a guarantee for every Apple OS, HDR transfer function or HEVC profile. The HEVC bitstream can still contain VUI color signaling after container `colr` is removed. Missing VUI or conflicting VUI/nclx signaling remains unqualified; retagging cannot repair an incorrect native matrix/range interpretation.

The source color space is never selected as the bitmap destination: retaining the source tag and drawing into sRGB would perform color matching before Core, causing duplicate ICC/gamut processing. Other native color models are rejected rather than silently converted into RGB.

## Resource limits

The constructor accepts a decoded-pixel limit (default 50,000,000) and encoded-byte limit (default 128 MiB), which callers can lower. Metadata dimensions are checked before native image creation and RGBA allocation. A separate 256 MiB working-set estimate accounts for the managed input, CFData copy, up to 64-bit native pixels, RGBA staging, the public raster API's copy and 64 KiB overhead. The working-set estimate is stricter than the pixel ceiling: for example, a single 18 MP tile is rejected even with a 50 MP pixel limit. A large grid made of small tiles can still be decoded by Core. Actual native dimensions and component depth are rechecked before bitmap allocation. ImageIO's private decoder scratch is opaque, so this estimate is not a hard process-memory limit. Owned Create/Copy objects, bitmap contexts, pins and explicitly loaded framework handles are released on success and failure.

## Native tests and reference helper

Run the dedicated tests on macOS; native cases skip on other operating systems:

```sh
MSBUILDDISABLENODEREUSE=1 DOTNET_CLI_USE_MSBUILD_SERVER=0 \
  dotnet test OfficeIMO.Drawing.ImageIO.Tests/OfficeIMO.Drawing.ImageIO.Tests.csproj \
  -c Release -f net8.0 -m:1 -nodeReuse:false -p:UseSharedCompilation=false
```

`OfficeIMO.Drawing.ImageIO.Tests/AppleImageIOReference.cs` is a test-only file that can be linked into verification programs. `DecodeFullFile` uses the native primary image with orientation/HEIF transforms and an sRGB destination, independently of Core composition and this codec. It forces decoding from the original image rather than an embedded preview. `DescribeStoredImage` records the native source space, depth and bitmap layout before drawing. `CreateColorfulHeic` encodes an original 32×24 pattern through `CGImageDestination`; `WriteColorfulHeic(path)` writes it without overwriting an existing file. Generated originals contain no camera metadata and may be retained as MIT fixtures; encoded bytes can vary across OS versions. The helper uses only Apple frameworks and is excluded from the production package.

`RealPhotoPickupMatchesNativeFullFileReferenceWhenProvided` reads `.heic`, `.heif` and `.hif` photos from `OFFICEIMO_HEIC_REAL_PHOTOS` or `OfficeIMO.TestAssets/Images/Heic/RealPhotos`. It skips when there are no photos. Provide originals with identifying metadata removed. The bounded native reference accepts up to 32 MiB encoded input and 16 MP per photo; larger inputs report a budget failure. Comparisons require matching dimensions, mean opaque RGB and mean premultiplied RGB errors of at most 3/255, and mean alpha error of at most 1/255. Output includes maximum errors and coordinates of RGB differences above 16/255; passing mean tolerances does not imply a per-pixel match.

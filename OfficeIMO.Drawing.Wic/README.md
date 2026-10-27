# OfficeIMO.Drawing.Wic

An optional HEIC decoder for OfficeIMO using Microsoft's Windows Imaging Component
(WIC) HEIF decoder and the user's installed HEVC support. Targets .NET 8 and .NET 10.
The package contains no third-party image decoder or HEVC implementation and does
not download, install, or license Windows extensions.

## Decode and register

Reference `OfficeIMO.Drawing.Wic` alongside the OfficeIMO package used by your
application. Windows requires working HEIF Image Extensions and HEVC Video
Extensions installed through Windows or Microsoft Store. `IsSupported` checks the
operating system; `GetAvailability()` checks the installed decoder and performs an
actual HEVC pixel decode.

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Wic;

var codec = new OfficeWicCodec();
OfficeRasterCodecAvailability availability = codec.GetAvailability();
if (!availability.IsAvailable)
    throw new InvalidOperationException(availability.Message);

// Register a process-wide fallback. Explicit ImageCodec options take precedence.
OfficeWicCodec.Register();

// Or call the codec directly for a single-item HEIC image.
if (codec.TryDecode(File.ReadAllBytes("tile.heic"), "image/heic", out var image))
    Console.WriteLine($"{image!.Width} x {image.Height}");
```

`Register()` replaces `OfficeRasterImageCodecs.Platform` only on Windows. It does
nothing on other systems and does not require the extensions to be installed at
registration time. Availability checks are not cached, so an application can
check again after a user installs or repairs an extension.

| Availability reason | Meaning |
| --- | --- |
| `None` | Microsoft's HEIF decoder is enumerated and the embedded HEVC probe decodes to RGBA. |
| `NoPlatformCodec` | The host is not Windows, or WIC enumeration failed; see the message. |
| `WindowsHeifExtensionMissing` | Microsoft's HEIF decoder is not present in WIC's decoder enumeration. |
| `WindowsHevcExtensionMissing` | The HEIF decoder is registered, but the HEVC pixel probe failed. The message includes the HRESULT when available. A broken codec can also cause this result. |

## Pixel contract and limits

The codec accepts HEIF signatures and asks only Microsoft's HEIF decoder for its
first frame. It returns top-to-bottom, tightly packed, straight-alpha RGBA with
8 bits per channel. The content type is a hint; file bytes select the decoder.
Corrupt, unavailable, unsupported, and over-budget inputs return `false` and a
null raster. Memory exhaustion is not disguised as an unavailable extension.
Callers must not modify encoded bytes while a decode is running.

OfficeIMO's HEIF composition layer supplies single-item HEIC without container
orientation, crop, or color properties. Container composition, alpha auxiliaries,
orientation, and color profiles belong to that layer. Directly passing a complete
photo to this codec exposes the installed WIC decoder's first-frame behavior,
rather than OfficeIMO's composition contract. This package provides no encoder.

Only `IWICFormatConverter` is used to obtain RGBA; there is no `IWICColorTransform`,
requested sRGB profile, explicit gamma correction, orientation operation, or HDR
tone mapper. The installed HEVC decoder controls YUV-to-RGB reconstruction,
including any interpretation of bitstream VUI matrix/range information. Requesting
RGBA does not provide raw YUV samples or a guarantee that VUI is ignored. Color
behavior must be qualified with the installed Windows extensions before treating
it as evidence for a particular color pipeline.

Each call checks these bounds before allocating managed output pixels:

- Encoded input: at most 128 MiB.
- Native frame dimensions: positive, representable as a tightly packed RGBA
  buffer, and at most 50,000,000 pixels.
- Working-memory admission: at most 256 MiB, counting the input, two managed RGBA
  buffers, an additional 8 bytes per pixel for native decode/conversion storage,
  and 64 KiB overhead. The working-memory limit can reject an image below the
  pixel limit. `OfficeRasterImage.FromRgba32` copies its input buffer.

The working-memory check bounds package-controlled allocations and reserves a
native allowance. WIC exposes no hard allocation quota for its installed decoder;
this is not a process-memory sandbox. WIC can allocate native metadata before
reporting dimensions. Untrusted files requiring a hard process-memory ceiling
need isolation provided by the host.

COM declarations use `GeneratedComInterface`, `LibraryImport`, and
`StrategyBasedComWrappers`, without runtime-generated COM stubs or `ComImport`.
Every call releases its unique wrappers and native references on its own thread
before balancing COM initialization. Existing STA apartments remain intact.
Input bytes remain pinned until all dependent native objects are released.
The codec is reentrant; returned mutable rasters retain their usual thread-safety
rules. .NET Framework 4.7.2 is not targeted because it lacks this generated COM
interop route.

## Availability probe

`Resources/Probe.heic` is a 16 by 16 solid RGB (96, 144, 192) image generated from a
synthetic PNG using macOS `sips -s format heic`, backed by system ImageIO. It is
original synthetic test data distributed under this package's MIT license,
contains no photograph or personal metadata, and is embedded only as encoded
image data. It is 507 bytes with SHA-256
`e97eea906fac7b8f81f3ab346bd0ea4c3db3eb335f64f073d5e37e175db0cdd9`.
The probe forces `CopyPixels`, so successful header parsing alone is insufficient.
It does not qualify every HEVC profile, bit depth, or device-produced photo.

## Windows validation

Run from the repository root with the pinned SDK. These commands do not install
or remove extensions. For NativeAOT, use a Windows machine with the supported
Visual C++ build tools and Windows SDK already installed.

```powershell
$env:MSBUILDDISABLENODEREUSE = '1'
$env:DOTNET_CLI_USE_MSBUILD_SERVER = '0'
$scratchRoot = if ($env:EVOTEC_SCRATCH_ROOT) { $env:EVOTEC_SCRATCH_ROOT } else { [IO.Path]::GetTempPath() }
$out = Join-Path $scratchRoot 'officeimo-wic-validation'

dotnet build OfficeIMO.Drawing.Wic/OfficeIMO.Drawing.Wic.csproj -c Release `
    --artifacts-path $out -p:UseSharedCompilation=false -nodeReuse:false

# With both extensions working, require actual availability (do not silently skip it).
$env:OFFICEIMO_WIC_EXPECT_AVAILABILITY = 'None'
dotnet test OfficeIMO.Drawing.Wic.Tests/OfficeIMO.Drawing.Wic.Tests.csproj `
    -c Release --artifacts-path $out -p:UseSharedCompilation=false -nodeReuse:false `
    --logger 'console;verbosity=normal'

dotnet publish OfficeIMO.Drawing.Wic.Tests/AotSmoke/OfficeIMO.Drawing.Wic.AotSmoke.csproj `
    -c Release -f net10.0 -r win-x64 `
    --artifacts-path $out -o (Join-Path $out 'native-smoke') `
    -p:UseSharedCompilation=false -nodeReuse:false
& (Join-Path $out 'native-smoke/OfficeIMO.Drawing.Wic.AotSmoke.exe') None
```

For ARM64 use `win-arm64`; for .NET 8 repeat the publish with `-f net8.0` and a
separate output folder. Unit tests skip on other operating systems. Decode tests
skip when the required extensions are unavailable, with the availability message
recorded as the skip reason.

On separate disposable Windows environments with known extension states, set
`OFFICEIMO_WIC_EXPECT_AVAILABILITY` to `WindowsHeifExtensionMissing` or
`WindowsHevcExtensionMissing` and repeat the tests. Pass that same reason to the
native smoke executable to require it there. The variable is used only by the
validation programs; it never alters codec selection or production behavior.
Without an expected reason, the native smoke executable fails when decoding is
unavailable on Windows.

For the three extension states, run the matching command on a Windows environment
already in that state. These commands assert the observed reason and never alter
the installed extensions:

```powershell
$env:MSBUILDDISABLENODEREUSE = '1'
$env:DOTNET_CLI_USE_MSBUILD_SERVER = '0'
$scratchRoot = if ($env:EVOTEC_SCRATCH_ROOT) { $env:EVOTEC_SCRATCH_ROOT } else { [IO.Path]::GetTempPath() }
$out = Join-Path $scratchRoot 'officeimo-wic-validation'
$tests = 'OfficeIMO.Drawing.Wic.Tests/OfficeIMO.Drawing.Wic.Tests.csproj'
$availability = 'FullyQualifiedName~Availability_ReportsRegisteredDecoderAndRealPixelDecode'

# HEIF decoder absent.
$env:OFFICEIMO_WIC_EXPECT_AVAILABILITY = 'WindowsHeifExtensionMissing'
dotnet test $tests -c Release -f net10.0 --filter $availability `
    --artifacts-path $out -p:UseSharedCompilation=false -nodeReuse:false

# HEIF registered; HEVC pixel decoding unavailable.
$env:OFFICEIMO_WIC_EXPECT_AVAILABILITY = 'WindowsHevcExtensionMissing'
dotnet test $tests -c Release -f net10.0 --filter $availability `
    --artifacts-path $out -p:UseSharedCompilation=false -nodeReuse:false

# Both extensions installed and functional: run every test.
$env:OFFICEIMO_WIC_EXPECT_AVAILABILITY = 'None'
dotnet test $tests -c Release --artifacts-path $out `
    -p:UseSharedCompilation=false -nodeReuse:false --logger 'console;verbosity=normal'
```

## Real camera photographs

The public CC0 corpus test requires matching dimensions and colors against the
native full-file reference for all eight fixtures when the extensions are
available. A missing reference fails that corpus check; dimension-only decoding
does not qualify the color path. Registration tests run in a collection with
parallel execution disabled because they change the process-wide fallback.

`RealPhotoPickupMatchesNativeFullFileReferenceWhenProvided` recursively discovers
`.heic`, `.heif`, and `.hif` under `OFFICEIMO_HEIC_REAL_PHOTOS`, or under
`OfficeIMO.TestAssets/Images/Heic/RealPhotos` when the variable is unset. It skips
on other operating systems, missing extensions, missing photos, or when no photo
has an available bounded WIC full-file reference. The test output lists each
unavailable reference; partial coverage is reported as a compared/total count.
Keep identifying originals outside the repository and provide GPS-stripped copies.

```powershell
$env:OFFICEIMO_HEIC_REAL_PHOTOS = 'D:\PrivateTestPhotos\Heic'
$env:OFFICEIMO_WIC_EXPECT_AVAILABILITY = 'None'
dotnet test OfficeIMO.Drawing.Wic.Tests/OfficeIMO.Drawing.Wic.Tests.csproj `
    -c Release -f net10.0 --artifacts-path $out `
    -p:UseSharedCompilation=false -nodeReuse:false `
    --filter 'FullyQualifiedName~RealPhotoPickup' --logger 'console;verbosity=detailed'
```

The test-only reference receives original full-file bytes directly through WIC,
reads WIC's HEIF orientation, and uses WIC color contexts and `IWICColorTransform`
for sRGB output. It shares only native ABI declarations and reference ownership
with the production adapter; it does not use Core parsing, tile synthesis,
composition, orientation, or color conversion. Integer RGB without a native
color context follows WIC's documented sRGB inference. The production tile codec
never uses this color-managed reference pipeline.

The reference accepts at most 32 MiB encoded input and 16 MP, subject to a 256 MiB
working-memory admission check including the retained Core result, reference
pixels, and native allowance. Comparisons avoid whole-image pixel snapshots.
Matching dimensions, mean opaque RGB and mean premultiplied RGB errors of at most
3/255, and mean alpha error of at most 1/255 are required. Maximum errors and the
first eight opaque RGB differences above 16/255 are logged; a passing mean does
not imply a per-pixel match. This is empirical qualification of the installed
Windows decoder, not a guarantee about every HEVC profile or internal VUI/color
decision.

Microsoft references: [generated COM interop](https://learn.microsoft.com/dotnet/standard/native-interop/comwrappers-source-generation),
[WIC decoder enumeration](https://learn.microsoft.com/windows/win32/api/wincodec/nf-wincodec-iwicimagingfactory-createcomponentenumerator),
[pinned memory lifetime](https://learn.microsoft.com/windows/win32/api/wincodec/nf-wincodec-iwicstream-initializefrommemory),
and [WIC pixel formats](https://learn.microsoft.com/windows/win32/wic/-wic-codec-native-pixel-formats).

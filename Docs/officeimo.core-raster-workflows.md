# Managed raster workflows

`OfficeIMO.Core` supplies image buffers, codecs, editing, metadata, comparison, and perceptual fingerprints through the `OfficeIMO.Drawing` namespace. These APIs use managed code and add no external runtime dependency. Browser capture and animation export belong to their respective rendering owners; a consumer can pass decoded Core frames to another renderer without exposing third-party image types in its own API.

## Read complete image sequences

Use `ReadEncodedBytes` when a workflow needs both decoded pixels and the original encoded metadata. It reads from the current stream position, returns owned bytes, leaves the stream open, and restores a seekable stream's original position on success or failure. A nonseekable stream is consumed and may advance one byte beyond the configured limit when detecting excess input.

```csharp
using System;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;

CancellationToken cancellationToken = default;
var options = new OfficeRasterDecodeOptions {
    MaximumEncodedBytes = 16 * 1024 * 1024,
    MaximumDecodedPixels = 20_000_000,
    CancellationToken = cancellationToken
};
using var input = File.OpenRead("input.gif");
byte[] bytes = OfficeRasterImageDecoder.ReadEncodedBytes(input, options);
OfficeRasterFrames frames = OfficeRasterImageDecoder.DecodeFrames(bytes, options, out var source);

OfficeRasterFrames grayscale = frames.Transform(
    image => OfficeRasterFilters.Grayscale(image, cancellationToken: cancellationToken),
    cancellationToken: cancellationToken);
byte[] firstFramePng = OfficeRasterImageEncoder.Encode(grayscale[0].Image, OfficeImageExportFormat.Png);
```

`TryDecodeFrames` returns every supported page or rendered animation frame in independent buffers. It fails without publishing a partial sequence if a frame is unsupported, incomplete, or over budget. GIF and APNG frames contain the composed canvas after blending and before disposal. Durations retain their source values, including zero; an animation writer must choose how its output format represents them. `PlayCount` expresses total plays: one means once, and zero means infinite. TIFF pages use one play and zero duration.

The eager decoder defaults to 256 frames and accepts an explicit maximum up to 4096. `MaximumDecodedPixels` applies to the entire decoded sequence. Encoded input and retained decoder working memory also remain bounded. Use `Decode` when only one page or frame is needed. The throwing methods raise `InvalidDataException` when decoding fails; the `TryDecode` and `TryDecodeFrames` variants retain best-effort workflows. Evidence overloads describe the detected container, decoded selection, and normalization independently of filenames.

JPEG and TIFF decoding normally apply EXIF display orientation. Set `ApplyExifOrientation = false` when an editing pipeline retains the original metadata and calls `OfficeRasterTransforms.AutoOrient` itself. After that explicit operation, replace the orientation tag with one before saving. PNG and WebP's direct managed decoders return stored sample order; this option preserves that existing behavior. The PNG conversion helper also honors the option for its WebP orientation step.

`OfficeRasterImageFormats.GetCapabilities(format)` describes the built-in inspection, pixel, multiple-frame, encoding, and metadata routes for an `OfficeImageFormat`. `BuiltInFormats` includes identification-only formats too. A route means the documented subset is supported; it does not promise every variant of that container. Caller-supplied codecs do not change these built-in descriptors. Use `OfficeImageExportFormat.GetContainerFormat()` to map an output choice to its container family.

## Edit images and frames

`OfficeRasterImage.FromRgba32` copies tightly packed top-to-bottom RGBA samples. `Clone` and the image filters return separately owned buffers. Geometric transforms return new images; `OfficeRasterCanvas` and `OfficeRasterText.Draw` deliberately draw into an existing image.

```csharp
OfficeRasterImage original = grayscale[0].Image;
OfficeRasterImage blurred = OfficeRasterFilters.GaussianBlur(original, sigma: 2);
OfficeRasterImage thumbnail = OfficeRasterResampler.Resize(blurred, 320, 200,
    OfficeRasterResamplingMode.Bicubic, OfficeRasterResamplingColorSpace.EncodedSrgb);
OfficeRasterText.Draw(thumbnail, "Preview", 12, 12, 296, 48, OfficeColor.White,
    new OfficeRasterTextOptions {
        FontSize = 24,
        ShadowColor = OfficeColor.Black,
        ShadowOffsetX = 1,
        ShadowOffsetY = 1
    }, cancellationToken);
```

Color amounts have explicit units: brightness, contrast, saturation, and lightness use one as identity; opacity and threshold use zero through one; hue uses degrees; blur radii and sigma use pixels. Grayscale supports BT.709 and BT.601. `OfficeRasterColorMatrix` copies twenty row-major coefficients and transforms normalized RGBA plus a constant bias. Neighborhood filters use premultiplied alpha to prevent hidden RGB from bleeding into visible pixels. Named photographic presets are Core's own visual effects; they do not promise byte equality with another library's implementation.

Resampling modes include nearest, bilinear, area, bicubic, box, triangle, Hermite, Lanczos kernels, Mitchell-Netravali, Robidoux variants, spline, and Welch. Choose the kernel and encoded-sRGB or linear-sRGB color space explicitly when they affect the result. Bicubic uses the Catmull-Rom kernel. The original four enum values retain their numeric identities.

## Resize to bounds or fill a target

`OfficeRasterResizeOptions` uses the same `OfficeImageFit` vocabulary as document image placement. `Contain` produces fitted pixels without padding and can infer one omitted dimension. `Cover` requires both dimensions and produces a centered crop. `Stretch` changes each supplied axis independently and keeps an omitted source axis. Editing options default to bicubic sampling; the existing exact-size overload retains its bilinear default.

```csharp
var resize = new OfficeRasterResizeOptions {
    Width = 320,
    Height = 200,
    Fit = OfficeImageFit.Cover,
    ResamplingMode = OfficeRasterResamplingMode.Bicubic
};
OfficeRasterResizePlan plan = OfficeRasterResampler.PlanResize(original.Width, original.Height,
    resize, cancellationToken);
OfficeRasterImage cover = OfficeRasterResampler.Resize(original, plan, cancellationToken);
OfficeRasterFrames covers = grayscale.Resize(resize, cancellationToken);
```

The immutable plan captures the output dimensions, intermediate dimensions, crop, sampling settings, and peak managed storage before pixel work begins. Changing the options later does not change that plan. Contain rounds to the nearest pixel, with midpoint ties away from zero, and remains inside the requested bounds. Cover rounds its intermediate dimensions upward; an odd extra pixel is cropped from the right or bottom. A 400 by 200 source fits a 100 by 100 contain target as 100 by 50 pixels. Cover samples it to 200 by 100 and crops the central 100 by 100.

`OfficeRasterFrames.Resize` plans the complete result and retained temporary storage before transforming any frame. It preserves timing and play count and publishes a complete new sequence. Oversized intermediates are rejected even when the final crop would be small.

## Pixel ownership and drawing

The constructor, `FromRgba32`, `Clone`, and `GetPixels` apply Core's pixel and managed-storage limits before allocating their buffers. `GetPixels` returns a separately owned RGBA snapshot; it has the same source-plus-copy budget as `Clone`. A valid large raster can therefore be readable while an additional full copy is rejected. Keep separately retained images in the caller's aggregate memory policy.

`GetPixel` requires valid coordinates. `SetPixel` and `BlendPixel` clip writes outside the image, supporting drawing at its edges. Canvas image drawing snapshots an aliased source before changing destination pixels, so drawing an image onto itself observes the original source samples. Its source copy is included in the drawing budget.

## Measure and draw with the same text settings

Pass `OfficeRasterTextOptions` to both `Measure` and `Draw` when text uses explicit fonts, shaping, wrapping, alignment, shadow, or outline effects. The operations use the existing canvas layout and shaping engine. Font size and line height must be positive and finite; a supplied measurement wrap width must also be positive and finite.

```csharp
var registeredFonts = new OfficeFontFaceCollection()
    .Add("Inter", File.ReadAllBytes("Inter-Regular.ttf"));
var caption = new OfficeRasterTextOptions {
    FontSize = 24,
    FontFamily = "Inter",
    Fonts = registeredFonts,
    Wrap = true,
    Clip = true
};
var measured = OfficeRasterText.Measure("Quarterly results", caption,
    wrapWidth: 280, cancellationToken: cancellationToken);
OfficeRasterText.Draw(cover, "Quarterly results", 20, 20, 280, 120,
    OfficeColor.White, caption, cancellationToken);
```

`registeredFonts` is the caller's `OfficeFontFaceCollection`, populated with the faces the workflow requires. Supplying the same collection to measurement and drawing avoids relying on each machine's installed fonts. Each operation snapshots the scalar options and font collection; font objects, shaping providers, and a diagnostic sink are retained references. Drawing changes the destination pixels as it runs. Cancellation stops further painting and does not undo pixels already drawn.

When composing text drawing with other retained buffers, use `OfficeRasterText.EstimateAdditionalWorkingBytes(width, height, options)` before allocating. It counts the mask and enabled outline storage beyond the destination raster, with row and fixed allowances. `Draw` uses the same calculation. `OfficeRasterTextOptions.Clone()` captures settings and its font collection for a complete caller workflow; providers and the diagnostic sink remain retained references.

`OfficeRasterFrames.Transform` accounts for the complete source and result sequences. Its `additionalRetainedBytes` argument reserves auxiliary images or peak temporary storage used inside the callback. For text, reserve the largest frame's text estimate and draw with the captured settings. Core also supplies `OfficeRasterFilters.EstimateGaussianBlurAdditionalWorkingBytes`, `EstimateBoxBlurAdditionalWorkingBytes`, `EstimateGaussianSharpenAdditionalWorkingBytes`, and `EstimateAdaptiveThresholdAdditionalWorkingBytes` for the filters with image-sized intermediates. Pass the same dimensions and settings used by the operation. The named estimates validate supported settings and include conservative kernel and fixed allowances; they describe managed operation storage, not total process memory.

## Encode with metadata and inspect omissions

`OfficeImageMetadata.Resolution` holds one immutable native-resolution value. The scalar resolution properties and generic EXIF density edits update that same state. Native units include physical density and a unitless aspect ratio. Profile imports and clones retain this authority; explicitly removing an EXIF density tag does not silently add that tag back during a rewrite.

```csharp
OfficeImageMetadata metadata = OfficeImageMetadata.Read(bytes, cancellationToken);
metadata.Resolution = new OfficeImageResolution(300, 300);
metadata.SetExifValue(OfficeExifTag.Software, "Photo workflow");

OfficeRasterEncodingResult encoded = OfficeRasterImageEncoder.EncodeWithMetadata(cover,
    OfficeImageExportFormat.Jpeg, metadata,
    new OfficeRasterEncodingOptions {
        Jpeg = new OfficeJpegEncodeOptions { Quality = 90 }
    }, cancellationToken: cancellationToken);
byte[] output = encoded.RequireMetadataPreservation();
OfficeImageFileWriter.WriteAllBytes("cover.jpg", output, cancellationToken);
```

The result reports supplied metadata profile families omitted by the emitted container. `RequireMetadataPreservation()` rejects those omissions; it does not promise lossless JPEG or WebP pixels. `EncodedBytes` is the independently owned output array. `Metadata` is an independent snapshot of the requested destination projection; read the encoded bytes to inspect exact density after native quantization.

The frame overload supports TIFF pages and icon resolutions, or a single still frame. It rejects multiple-frame output in still formats that cannot retain it. Metadata edits describe the primary TIFF image, with a shared page-density setting; they do not promise per-page EXIF or profile preservation. GIF/APNG pixel encoding stays with the animation owner. Its completed bytes can pass through `OfficeImageMetadata.ApplyForEncoding` for the same metadata projection and omission evidence.

`OfficeRasterEncodingOptions.Resolution` is an explicit nullable override. In the pixel encoder, null leaves the selected codec's own density settings in control. In metadata-aware encoding, null takes density from the supplied metadata, while a nonnull value overrides it. `WriteResolutionMetadata` controls density emission. This separates an intentional override from reading and reassigning a default scalar value.

`OfficeRasterFrames.Transform` preserves duration and playback count while returning a complete new sequence. Supply an output-size planner for operations that change dimensions. Planning finishes before the first mapping call:

```csharp
OfficeRasterFrames rotated = grayscale.Transform(
    image => OfficeRasterTransforms.Rotate(image, 90, cancellationToken: cancellationToken),
    image => OfficeRasterTransforms.GetRotatedSize(image, 90),
    cancellationToken);
```

The source and planned result buffers together must fit the collection's retained-memory budget. Arbitrary mapping callbacks own their temporary-memory policy and should leave source pixels unchanged. This is a per-operation and retained-buffer policy, rather than a hard cap on total process memory. Keep independently retained copies and downstream renderer buffers in the caller's aggregate budget.

## Compare pixels and compute visual fingerprints

`OfficeRasterComparison.Compare` requires equal dimensions. It compares premultiplied RGB and alpha, so completely transparent pixels with different hidden RGB values compare equal. The result includes changed-pixel count, the largest channel difference, a normalized mean absolute difference, and an opaque difference image. `Similarity` is one minus that normalized mean; it is a pixel metric rather than a perceptual confidence score.

```csharp
OfficeRasterComparisonResult comparison = OfficeRasterComparison.Compare(original, blurred);
byte[] differencePng = OfficeRasterImageEncoder.Encode(comparison.DifferenceImage,
    OfficeImageExportFormat.Png);
ulong hash = OfficeRasterFingerprinting.DifferenceHash(original);
string fingerprint = hash.ToString("x16", System.Globalization.CultureInfo.InvariantCulture);
```

The 64-bit difference hash resizes to 9 by 8 with bicubic filtering in encoded sRGB, converts to rounded BT.709 luminance, and sets a bit when a pixel is brighter than its right neighbor. Bits run left to right and top to bottom, beginning at the least significant bit. This is a visual similarity signal. Persisted fingerprints must be rebuilt when decoding or resampling implementations change; compare hashes generated by the same policy.

## Inspect and replace HEIF metadata

Use `OfficeHeifMetadataReader` for HEIF/HEIC container information and independent EXIF or
XMP edits. This API works with encoded container bytes and preserves image payloads; it does
not decode or encode HEIF pixels.

```csharp
byte[] heif = File.ReadAllBytes("photo.heic");
if (OfficeHeifMetadataReader.TryReadExifProfile(heif, out OfficeImageMetadata? exif,
        cancellationToken) && exif != null) {
    exif.SetExifValue(OfficeExifTag.Software, "Photo workflow");
    if (OfficeHeifMetadataReader.TryWriteExifProfile(heif, exif, out byte[]? edited,
            cancellationToken)) {
        OfficeImageFileWriter.WriteAllBytes("annotated.heic", edited!, cancellationToken);
    }
}
```

`TryReadInfo` returns brands, primary image properties, item associations, references, and
locations through the `OfficeHeif*` models. `HasExifItem` and `HasXmpItem` distinguish a
declared but unlocated item from a missing family. Reads and writes select the unique metadata
item whose `cdsc` reference describes the `pitm` primary image, independently of declaration
order. A sole unassociated item remains supported for older containers. Ambiguous items or
items explicitly associated only with another image are not selected. Information retains
all declared items, with `ExifItem` and `XmpItem` identifying the selection or returning null.
Incomplete item, reference, location, and property-association collections return `false`
without partial information. Protected items and XMP items declaring
a MIME content encoding are visible as opaque declarations; payload reads and writes,
including clearing, return `false`. An independent edit to another family preserves those
opaque bytes. Readers support absolute file extents,
item-data-box extents, and multiple extents within the bounded parser contract. Stream readers
leave the stream open and restore its original position when seeking is available.

Writing requires an existing single absolute extent inside an `mdat` payload. It rejects
shared extents before erasing old bytes, leaves unrequested metadata families unchanged, and
places replacement data in a framed `mdat` box. Pass null to clear an existing profile;
missing items are not created. Byte-array writers return owned bytes and never mutate the
source. File writers prepare the complete output first, so rejection and cancellation during
preparation leave an existing output file untouched. File writers stage complete output beside
the destination and atomically commit it, preserving the existing file if a staging write fails.

The policy bounds encoded input/output to 128 MiB, individual metadata and property payloads
to 16 MiB, declared item collections to 4,096 entries, and parser work to 65,536 records.
Parser collections, copied payloads, retained EXIF input, known caller stream backing, and prepared output are accounted
against Core's 256 MiB managed working-set limit. Cancellation is observed during structure
scanning, payload copies, and rewrite preparation.

XMP reads reject malformed UTF-8, and writes reject unpaired UTF-16 surrogates instead of
replacing invalid text. Independent EXIF edits preserve malformed XMP bytes. A valid new
packet can replace malformed XMP through the direct writer.
The writer includes the caller's UTF-16 packet storage in its managed working-set budget.

## Save complete encoded output

`OfficeImageFileWriter` saves completed encoded bytes through Core's shared atomic file
commit owner. It creates missing parent directories, stages in the destination directory,
and atomically creates or replaces the output:

```csharp
OfficeImageFileWriter.WriteAllBytes("preview.png", firstFramePng, cancellationToken);
await OfficeImageFileWriter.WriteAllBytesAsync("difference.png", differencePng, cancellationToken);
```

Both methods borrow the supplied byte array; keep it unchanged until completion. They do
not clone, parse, or impose format limits on already encoded data. Failed staging and
cancellation observed before commit preserve an existing destination, and failed staging
files are cleaned up when the filesystem permits it. Unsupported atomic replacement fails
explicitly. The final replacement can complete if cancellation arrives after commit begins.
Atomic publication protects readers from partial output; it does not promise power-loss durability.

## Failure and cancellation

`TryDecode` and `TryDecodeFrames` return `false` for unsupported, malformed, or over-budget input. `Decode` and `DecodeFrames` throw `InvalidDataException` when decoding fails. Both forms propagate cancellation. Editing and encoding APIs reject unsupported dimensions, parameters, or resource requirements with an exception. Expensive pixel operations observe their cancellation token during work. A failed collection transformation returns no partial result, although a caller-supplied callback can still mutate the pixels it was given. Encoding to a caller-owned stream leaves it open and can leave partial output after cancellation or an I/O failure.

Core does not provide a native-code fallback for unsupported formats or modes. Metadata editing, format conversion, and frame-loss policy are explicit contracts; consult the Core README and each codec's options before changing an image's container.

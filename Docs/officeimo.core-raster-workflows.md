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
if (!OfficeRasterImageDecoder.TryDecodeFrames(bytes, options, out var frames)) {
    throw new InvalidDataException("The complete image sequence cannot be decoded within this policy.");
}

OfficeRasterFrames grayscale = frames!.Transform(
    image => OfficeRasterFilters.Grayscale(image, cancellationToken: cancellationToken),
    cancellationToken: cancellationToken);
byte[] firstFramePng = OfficeRasterImageEncoder.Encode(grayscale[0].Image, OfficeImageExportFormat.Png);
```

`TryDecodeFrames` returns every supported page or rendered animation frame in independent buffers. It fails without publishing a partial sequence if a frame is unsupported, incomplete, or over budget. GIF and APNG frames contain the composed canvas after blending and before disposal. Durations retain their source values, including zero; an animation writer must choose how its output format represents them. `PlayCount` expresses total plays: one means once, and zero means infinite. TIFF pages use one play and zero duration.

The eager decoder defaults to 256 frames and accepts an explicit maximum up to 4096. `MaximumDecodedPixels` applies to the entire decoded sequence. Encoded input and retained decoder working memory also remain bounded. Use the existing selected-frame decoder when only one page or frame is needed.

JPEG and TIFF decoding normally apply EXIF display orientation. Set `ApplyExifOrientation = false` when an editing pipeline retains the original metadata and calls `OfficeRasterTransforms.AutoOrient` itself. After that explicit operation, replace the orientation tag with one before saving. PNG and WebP's direct managed decoders return stored sample order; this option preserves that existing behavior. The PNG conversion helper also honors the option for its WebP orientation step.

## Edit images and frames

`OfficeRasterImage.FromRgba32` copies tightly packed top-to-bottom RGBA samples. `Clone` and the image filters return separately owned buffers. Geometric transforms return new images; `OfficeRasterCanvas` and `OfficeRasterText.Draw` deliberately draw into an existing image.

```csharp
OfficeRasterImage original = grayscale[0].Image;
OfficeRasterImage blurred = OfficeRasterFilters.GaussianBlur(original, sigma: 2);
OfficeRasterImage thumbnail = OfficeRasterResampler.Resize(blurred, 320, 200,
    OfficeRasterResamplingMode.Bicubic, OfficeRasterResamplingColorSpace.EncodedSrgb);
OfficeRasterText.Draw(thumbnail, "Preview", 12, 12, 296, 48, OfficeColor.White,
    fontSize: 24, shadowColor: OfficeColor.Black, shadowOffsetX: 1, shadowOffsetY: 1);
```

Color amounts have explicit units: brightness, contrast, saturation, and lightness use one as identity; opacity and threshold use zero through one; hue uses degrees; blur radii and sigma use pixels. Grayscale supports BT.709 and BT.601. `OfficeRasterColorMatrix` copies twenty row-major coefficients and transforms normalized RGBA plus a constant bias. Neighborhood filters use premultiplied alpha to prevent hidden RGB from bleeding into visible pixels. Named photographic presets are Core's own visual effects; they do not promise byte equality with another library's implementation.

Resampling modes include nearest, bilinear, area, bicubic, box, triangle, Hermite, Lanczos kernels, Mitchell-Netravali, Robidoux variants, spline, and Welch. Choose the kernel and encoded-sRGB or linear-sRGB color space explicitly when they affect the result. Bicubic uses the Catmull-Rom kernel. The original four enum values retain their numeric identities.

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
declared but unlocated item from a missing family. Protected items and XMP items declaring
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

Decode methods return `false` for unsupported, malformed, or over-budget input and propagate cancellation. Editing and encoding APIs reject unsupported dimensions, parameters, or resource requirements with an exception. Expensive pixel operations observe their cancellation token during work. A failed collection transformation returns no partial result, although a caller-supplied callback can still mutate the pixels it was given. Encoding to a caller-owned stream leaves it open and can leave partial output after cancellation or an I/O failure.

Core does not provide a native-code fallback for unsupported formats or modes. Metadata editing, format conversion, and frame-loss policy are explicit contracts; consult the Core README and each codec's options before changing an image's container.

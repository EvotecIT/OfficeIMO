using OfficeIMO.Drawing;
using OfficeIMO.Epub;
using OfficeIMO.Provenance;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    /// <summary>Local KDP handoff ZIP ceiling, including the publication and separate cover.</summary>
    public const long MaximumKdpDeliveryBytes = 184L * 1024 * 1024;
    /// <summary>Conservative local cover byte cap: strictly below 50 decimal MB.</summary>
    public const int MaximumKdpCoverBytes = 49_999_999;
    /// <summary>Local cover decoding limit, independent of the retailer's dimension rules.</summary>
    public const long MaximumKdpCoverPixels = 16_000_000;

    /// <summary>
    /// Prepares an offline KDP handoff containing publication.epub, package.opf, a separate
    /// listing cover and a versioned manifest. The ZIP is an OfficeIMO handoff, not an upload format.
    /// </summary>
    /// <remarks>
    /// Requires a single JPEG or TIFF supported by the managed decoder, at least 625 by 1000 pixels,
    /// at most 10000 pixels per axis and within the local byte/pixel bounds. Dimensions describe
    /// the decoded display after embedded orientation is applied. Cover bytes are preserved.
    /// Color mode, color separation, visual orientation, quality, listing consistency, rights,
    /// accessibility, Kindle Previewer and retailer acceptance remain explicit unchecked scopes.
    /// The EPUB uses the same import-review, signature and writer policies as Export.
    /// No files are written, accounts accessed or publications uploaded.
    /// </remarks>
    public byte[] ToKdpDeliveryBytes(byte[] listingCover, EpubWriteOptions? options = null,
        long maximumOutputBytes = MaximumKdpDeliveryBytes, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(listingCover);
        if (maximumOutputBytes <= 0 || maximumOutputBytes > MaximumKdpDeliveryBytes)
            throw new ArgumentOutOfRangeException(nameof(maximumOutputBytes));
        cancellationToken.ThrowIfCancellationRequested();
        if (listingCover.Length == 0 || listingCover.Length > MaximumKdpCoverBytes)
            throw new InvalidDataException("The listing cover must be nonempty and smaller than 50,000,000 bytes.");
        // Keep the checked cover, its hash and its packaged bytes bound to one snapshot.
        byte[] cover = (byte[])listingCover.Clone();
        OfficeImageInfo info = ValidateKdpCover(cover, cancellationToken);
        var (publication, package, delivery) = CreateDeliveryPayloads(options, maximumOutputBytes, cancellationToken);
        string coverName = info.Format == OfficeImageFormat.Jpeg ? "listing-cover.jpeg" : "listing-cover.tiff";
        delivery.Files = [.. delivery.Files, DescribeDeliveryFile(coverName, info.MimeType, cover)];
        var manifest = new BookKdpDeliveryManifest {
            Delivery = delivery,
            Cover = new BookKdpCoverChecks {
                File = coverName, Width = info.Width, Height = info.Height,
                Recommendations = KdpCoverRecommendations(info)
            }
        };
        byte[] manifestBytes;
        using (var output = new OfficeProvenanceBoundedMemoryStream(MaximumReviewBytes)) {
            JsonSerializer.Serialize(output, manifest, BookKdpDeliveryJsonContext.Default.BookKdpDeliveryManifest);
            manifestBytes = output.ToArray();
        }
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeProvenanceZipWriter.Write([
            Entry("publication.epub", publication, false), Entry("package.opf", package, true),
            Entry(coverName, cover, false), Entry("manifest.json", manifestBytes, true)
        ], MaximumPublicationBytes + MaximumDeliveryMetadataBytes + MaximumReviewBytes + MaximumKdpCoverBytes,
            maximumOutputBytes: maximumOutputBytes, cancellationToken: cancellationToken);
    }

    private static OfficeImageInfo ValidateKdpCover(byte[] cover, CancellationToken cancellationToken) {
        if (!OfficeImageReader.TryIdentifyByContent(cover, null, out OfficeImageInfo info) ||
            info.Format is not (OfficeImageFormat.Jpeg or OfficeImageFormat.Tiff))
            throw new InvalidDataException("The listing cover must contain JPEG or TIFF image data.");
        if (info.Width > 10000 || info.Height > 10000)
            throw new InvalidDataException("The listing cover must be at most 10000 pixels per axis.");
        if ((long)info.Width * info.Height > MaximumKdpCoverPixels)
            throw new InvalidDataException("The listing cover exceeds the local 16,000,000-pixel decoding limit.");
        var decodeOptions = new OfficeRasterDecodeOptions {
            MaximumEncodedBytes = MaximumKdpCoverBytes, MaximumDecodedPixels = MaximumKdpCoverPixels,
            FrameLossPolicy = OfficeRasterFrameLossPolicy.RejectMultipleFrames,
            CancellationToken = cancellationToken
        };
        if (!OfficeRasterImageDecoder.TryDecode(cover, decodeOptions, out var image, out _) || image == null)
            throw new InvalidDataException("The listing cover is malformed, multi-page, or outside the managed decoder's supported subset or limits.");
        // JPEG header dimensions are stored axes; the decoder applies EXIF orientation.
        // TIFF identification already applies orientation. Use one display-space contract for both.
        if (image.Width < 625 || image.Height < 1000 || image.Width > 10000 || image.Height > 10000)
            throw new InvalidDataException("The listing cover must display at least 625 by 1000 pixels and at most 10000 pixels per axis after embedded orientation.");
        return new OfficeImageInfo(info.Format, image.Width, image.Height);
    }

    private static string[] KdpCoverRecommendations(OfficeImageInfo info) {
        var recommendations = new List<string>();
        if (info.Width < 1600 || info.Height < 2560)
            recommendations.Add("Consider a cover of at least 1600 by 2560 pixels for the recommended display quality.");
        if ((long)info.Height * 10 < (long)info.Width * 16)
            recommendations.Add("Consider the recommended height-to-width ratio of at least 1.6:1.");
        return recommendations.ToArray();
    }
}

internal sealed class BookKdpDeliveryManifest {
    public string Format { get; set; } = "OfficeIMO.KdpDelivery";
    public int Version { get; set; } = 1;
    public string Purpose { get; set; } = "offline-handoff-extract-files-before-upload";
    public string RequirementsSource { get; set; } = "https://kdp.amazon.com/en_US/help/topic/G200645690";
    public BookDeliveryManifest Delivery { get; set; } = new();
    public BookKdpCoverChecks Cover { get; set; } = new();
    public string KindlePreviewer { get; set; } = "not-performed";
    public string ListingMetadataRightsAndCommercialTerms { get; set; } = "not-checked";
    public string AccessibilityAssessment { get; set; } = "not-performed";
}

internal sealed class BookKdpCoverChecks {
    public string File { get; set; } = string.Empty;
    public int Width { get; set; }
    public int Height { get; set; }
    public string DimensionBasis { get; set; } = "decoded-display-after-embedded-orientation";
    public string FormatAndDimensions { get; set; } = "passed";
    public string ManagedPixelDecode { get; set; } = "passed";
    public long MaximumEncodedBytes { get; set; } = BookProject.MaximumKdpCoverBytes;
    public long MaximumDecodedPixels { get; set; } = BookProject.MaximumKdpCoverPixels;
    public string ColorModeAndSeparation { get; set; } = "not-checked";
    public string OrientationResolutionAndVisualQuality { get; set; } = "not-checked";
    public string ListingAndEmbeddedCoverConsistency { get; set; } = "not-checked";
    public string[] Recommendations { get; set; } = [];
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(BookKdpDeliveryManifest))]
internal sealed partial class BookKdpDeliveryJsonContext : JsonSerializerContext;

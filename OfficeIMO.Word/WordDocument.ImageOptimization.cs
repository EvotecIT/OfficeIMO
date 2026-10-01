using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>
    /// Evaluates exact encoded-media candidates without changing the document or fetching linked images.
    /// Candidate byte savings do not include ZIP packaging and document-save normalization.
    /// </summary>
    public WordImageOptimizationReport AnalyzeImageOptimization(WordImageOptimizationOptions? options = null,
        CancellationToken cancellationToken = default) => OptimizeImagesCore(options, apply: false, cancellationToken);

    /// <summary>
    /// Optimizes unique embedded raster media while retaining every relationship and drawing's geometry.
    /// All candidates are staged before mutation. Cancellation or candidate failure leaves media unchanged.
    /// Save through the normal Word save API to persist changes; signed packages are blocked by default.
    /// </summary>
    public WordImageOptimizationReport OptimizeImages(WordImageOptimizationOptions? options = null,
        CancellationToken cancellationToken = default) => OptimizeImagesCore(options, apply: true, cancellationToken);

    private WordImageOptimizationReport OptimizeImagesCore(WordImageOptimizationOptions? options, bool apply,
        CancellationToken token) {
        token.ThrowIfCancellationRequested();
        WordImageOptimizationOptions policy = (options ?? new WordImageOptimizationOptions()).Clone();
        if (apply) {
            if (FileOpenAccess == FileAccess.Read) throw new InvalidOperationException("A read-only document cannot be optimized.");
            EnsureSignedDocumentSaveAllowed(new WordSaveOptions { SignedDocumentPolicy = policy.SignedDocumentPolicy }, "OptimizeImages");
        }
        var inventory = WordImageOptimizationInventory.Build(_wordprocessingDocument, policy, token);
        var items = new List<WordImageOptimizationItem>();
        var staged = new List<WordImagePartReplacement>();
        long retainedBytes = 0;
        foreach (var media in inventory.Images.Values.OrderBy(image => image.Part.Uri.ToString(), StringComparer.Ordinal)) {
            token.ThrowIfCancellationRequested();
            using Stream stream = media.Part.GetStream(FileMode.Open, FileAccess.Read);
            byte[] original = OfficeStreamReader.ReadAllBytes(stream, token,
                apply ? Math.Min(policy.MaxImageBytes, policy.MaxStagedBytes - retainedBytes) : policy.MaxImageBytes);
            var info = OfficeImageReader.TryIdentify(original, media.Part.Uri.ToString(), out OfficeImageInfo identified)
                ? identified : new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
            WordImageOptimizationStatus? preserve = media.References == 0 ? WordImageOptimizationStatus.Unreferenced : null;
            if (!IsWordOptimizationFormat(info.Format)) preserve = WordImageOptimizationStatus.UnsupportedFormat;
            else if (info.Width <= 0 || info.Height <= 0) preserve = WordImageOptimizationStatus.DecodeFailed;
            else if (policy.Mode != OfficeImageOptimizationMode.Recompress && media.UnknownPlacement)
                preserve ??= WordImageOptimizationStatus.UnknownPlacement;
            if (preserve.HasValue) {
                items.Add(new WordImageOptimizationItem(media.Part.Uri.ToString(), media.References,
                    preserve.Value, original.LongLength, original.LongLength, info, info));
                continue;
            }
            int sourceWidth = info.Width, sourceHeight = info.Height;
            if (OfficeImageOrientationNormalizer.TryRead(original, out OfficeImageOrientation orientation) && orientation >= OfficeImageOrientation.Transpose)
                (sourceWidth, sourceHeight) = (sourceHeight, sourceWidth);
            int targetWidth = sourceWidth, targetHeight = sourceHeight;
            if (policy.Mode != OfficeImageOptimizationMode.Recompress) {
                // Cover both axes rather than fitting within a box: stretched placements must not lose one axis's required resolution.
                double scale = Math.Min(1D, Math.Max(media.TargetWidth / (double)sourceWidth, media.TargetHeight / (double)sourceHeight));
                targetWidth = Math.Max(1, (int)Math.Ceiling(sourceWidth * scale));
                targetHeight = Math.Max(1, (int)Math.Ceiling(sourceHeight * scale));
            }
            OfficeImageOptimizationResult candidate = OfficeImageOptimizer.Optimize(original,
                new OfficeImageOptimizationRequest(targetWidth, targetHeight) {
                    Mode = policy.Mode, OutputFormat = info.Format is OfficeImageFormat.Bmp or OfficeImageFormat.Gif ? OfficeImageFormat.Png : info.Format,
                    CancellationToken = token,
                    JpegQuality = policy.JpegQuality, KeepOriginalWhenNotSmaller = policy.KeepOriginalWhenNotSmaller,
                    ResamplingMode = policy.ResamplingMode, MetadataPolicy = policy.MetadataPolicy,
                    MetadataSelection = policy.MetadataSelection
                }, media.Part.Uri.ToString());
            token.ThrowIfCancellationRequested();
            bool replace = candidate.Changed && (!candidate.Metadata.HasLoss || policy.AllowMetadataLoss);
            var status = candidate.Changed && !replace ? WordImageOptimizationStatus.MetadataLoss : MapImageOptimizationStatus(candidate.Status);
            items.Add(new WordImageOptimizationItem(media.Part.Uri.ToString(), media.References, status,
                original.LongLength, replace ? candidate.FinalEncodedLength : original.LongLength,
                info, replace ? candidate.Final : info, candidate.Metadata));
            if (replace) {
                byte[] bytes = candidate.Bytes;
                if (apply) {
                    retainedBytes = checked(retainedBytes + original.LongLength + bytes.LongLength);
                    if (retainedBytes > policy.MaxStagedBytes) throw new InvalidDataException("Image optimization exceeds the staged-byte limit.");
                }
                if (!OfficeImageReader.TryValidateContent(bytes, null, out _))
                    throw new InvalidDataException("The optimized image candidate failed content validation.");
                if (apply) staged.Add(new WordImagePartReplacement(media.Part, original, bytes,
                    candidate.Final.Format == info.Format ? media.Part.ContentType : candidate.Final.MimeType));
            }
        }
        token.ThrowIfCancellationRequested();
        if (apply) WordImagePartReplacement.Apply(_wordprocessingDocument.MainDocumentPart!, staged, token);
        return new WordImageOptimizationReport(items, inventory.ExternalReferences, apply);
    }

    private static bool IsWordOptimizationFormat(OfficeImageFormat format) =>
        format == OfficeImageFormat.Png || format == OfficeImageFormat.Jpeg ||
        format == OfficeImageFormat.Tiff || format == OfficeImageFormat.Webp ||
        format == OfficeImageFormat.Bmp || format == OfficeImageFormat.Gif;

    private static WordImageOptimizationStatus MapImageOptimizationStatus(OfficeImageOptimizationStatus status) => status switch {
        OfficeImageOptimizationStatus.Optimized => WordImageOptimizationStatus.Optimized,
        OfficeImageOptimizationStatus.AlreadySuitable => WordImageOptimizationStatus.AlreadySuitable,
        OfficeImageOptimizationStatus.OriginalWasSmaller => WordImageOptimizationStatus.OriginalWasSmaller,
        OfficeImageOptimizationStatus.UnsupportedFormat => WordImageOptimizationStatus.UnsupportedFormat,
        _ => WordImageOptimizationStatus.DecodeFailed
    };

}

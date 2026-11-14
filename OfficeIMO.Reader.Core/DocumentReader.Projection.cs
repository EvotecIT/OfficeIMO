namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    internal const string TextAndMetadataProjectionCapability = "officeimo.reader.projection.text-and-metadata";
    internal const string TextProjectionCapability = "officeimo.reader.projection.text";

    /// <summary>Whether the current read builds the lean profile, either for itself or as the source of a text read.</summary>
    internal static bool IsTextAndMetadataProjection =>
        ReaderReadScope.Current?.Projection is ReaderProjection.TextAndMetadata or ReaderProjection.Text;

    /// <summary>Whether the current read returns only text chunks, so a handler may read text without its document model.</summary>
    internal static bool IsTextProjection => ReaderReadScope.Current?.Projection == ReaderProjection.Text;

    /// <summary>The projection the caller asked for, including a text request carried on normalized options.</summary>
    internal static ReaderProjection RequestedProjection(ReaderOptions? options) =>
        options?.TextOnly == true ? ReaderProjection.Text : options?.Projection ?? ReaderProjection.Full;

    /// <summary>Nested and repeated reads keep the narrowest profile requested so far: Full, then TextAndMetadata, then Text.</summary>
    internal static ReaderProjection NarrowestProjection(ReaderProjection left, ReaderProjection right) =>
        ProjectionRank(left) >= ProjectionRank(right) ? left : right;

    private static int ProjectionRank(ReaderProjection projection) => projection switch {
        ReaderProjection.Text => 2,
        ReaderProjection.TextAndMetadata => 1,
        _ => 0
    };

    internal static bool IsProjectionEligible(ReaderHandlerDescriptor handler, ReaderOptions options, string? sourceName) =>
        options.Projection == ReaderProjection.Full ||
        handler.SupportsTextAndMetadata && (handler.TextAndMetadataExtensions.Count == 0 ||
            handler.TextAndMetadataExtensions.Contains(
                NormalizeExtension(TryGetExtension(sourceName ?? string.Empty)), StringComparer.OrdinalIgnoreCase));

    private static void ValidateProjection(ReaderHandlerDescriptor handler, ReaderOptions options, string? sourceName) {
        if (!IsProjectionEligible(handler, options, sourceName)) {
            throw new NotSupportedException($"Reader handler '{handler.Id}' does not support the TextAndMetadata projection.");
        }
    }

    internal static bool HasTextAndMetadataProjection(OfficeDocumentReadResult result) =>
        result.CapabilitiesUsed.Contains(TextAndMetadataProjectionCapability, StringComparer.Ordinal)
        || result.CapabilitiesUsed.Contains(TextProjectionCapability, StringComparer.Ordinal);

    internal static void ApplyProjectionCapability(OfficeDocumentReadResult result, ReaderProjection projection) {
        string capability;
        if (projection == ReaderProjection.TextAndMetadata) capability = TextAndMetadataProjectionCapability;
        else if (projection == ReaderProjection.Text) capability = TextProjectionCapability;
        else return;
        result.CapabilitiesUsed = result.CapabilitiesUsed.Concat(new[] { capability })
            .Distinct(StringComparer.Ordinal).ToArray();
    }
}

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    internal const string TextAndMetadataProjectionCapability = "officeimo.reader.projection.text-and-metadata";

    internal static bool IsTextAndMetadataProjection =>
        ReaderReadScope.Current?.Projection == ReaderProjection.TextAndMetadata;

    internal static bool IsProjectionEligible(ReaderHandlerDescriptor handler, ReaderOptions options, string? sourceName) =>
        options.Projection != ReaderProjection.TextAndMetadata ||
        handler.SupportsTextAndMetadata && (handler.TextAndMetadataExtensions.Count == 0 ||
            handler.TextAndMetadataExtensions.Contains(
                NormalizeExtension(TryGetExtension(sourceName ?? string.Empty)), StringComparer.OrdinalIgnoreCase));

    private static void ValidateProjection(ReaderHandlerDescriptor handler, ReaderOptions options, string? sourceName) {
        if (!IsProjectionEligible(handler, options, sourceName)) {
            throw new NotSupportedException($"Reader handler '{handler.Id}' does not support the TextAndMetadata projection.");
        }
    }

    internal static bool HasTextAndMetadataProjection(OfficeDocumentReadResult result) =>
        result.CapabilitiesUsed.Contains(TextAndMetadataProjectionCapability, StringComparer.Ordinal);

    internal static void ApplyProjectionCapability(OfficeDocumentReadResult result, ReaderProjection projection) {
        if (projection != ReaderProjection.TextAndMetadata) return;
        result.CapabilitiesUsed = result.CapabilitiesUsed.Concat(new[] { TextAndMetadataProjectionCapability })
            .Distinct(StringComparer.Ordinal).ToArray();
    }
}

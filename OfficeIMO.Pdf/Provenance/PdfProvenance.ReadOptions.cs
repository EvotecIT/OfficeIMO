using OfficeIMO.Provenance;

namespace OfficeIMO.Pdf;

public static partial class PdfProvenance {
    /// <summary>
    /// Combines provenance carrier limits with PDF reader limits. A caller-supplied reader limit remains
    /// authoritative; synthesized options inherit a raised provenance asset limit so both layers accept the same input.
    /// </summary>
    internal static PdfLoadOptions CreateReadOptionsForInspection(
        OfficeProvenanceOptions options,
        PdfLoadOptions? readOptions) {
        Guard.NotNull(options, nameof(options));
        OfficeProvenanceBinary.ValidateLimits(options);
        return CreateReadOptions(
            options.MaxAssetBytes,
            options.MaxContainerEntries,
            options.MaxExpandedContainerBytes,
            options.MaxManifestBytes,
            GetMaximumManifestBytes(options),
            readOptions);
    }

    private static PdfLoadOptions CreateReadOptions(
        long maximumAssetBytes,
        int maximumContainerEntries,
        long maximumExpandedContainerBytes,
        long maximumSingleManifestBytes,
        long maximumTotalManifestBytes,
        PdfLoadOptions? readOptions) {
        PdfLoadOptions effective = PdfLoadOptions.WithMaximumContainerEntries(
            readOptions,
            maximumContainerEntries,
            maximumDecodedStreamBytes: readOptions == null
                ? Math.Min(
                    maximumExpandedContainerBytes,
                    Math.Max(PdfReadLimits.Default.MaxDecodedStreamBytes, maximumSingleManifestBytes))
                : maximumExpandedContainerBytes,
            maximumTotalDecodedStreamBytes: readOptions == null
                ? maximumExpandedContainerBytes
                : Math.Min(readOptions.Limits.MaxTotalDecodedStreamBytes, maximumExpandedContainerBytes),
            maximumTotalAttachmentBytes: readOptions == null ? maximumTotalManifestBytes : null,
            preserveExistingDecodedStreamLimit: readOptions != null,
            maximumRawStreamBytes: readOptions == null
                ? Math.Max(PdfReadLimits.Default.MaxRawStreamBytes, maximumSingleManifestBytes)
                : null,
            preserveExistingRawStreamLimit: readOptions != null);
        return readOptions == null
            ? PdfLoadOptions.WithMinimumInputBytes(effective, maximumAssetBytes)
            : effective;
    }

    private static long GetMaximumManifestBytes(OfficeProvenanceOptions options) => Math.Min(
        options.MaxExpandedContainerBytes,
        MultiplySaturating(options.MaxManifestBytes, options.MaxCarriers));

    private static long MultiplySaturating(long value, int multiplier) =>
        value > long.MaxValue / multiplier ? long.MaxValue : value * multiplier;
}

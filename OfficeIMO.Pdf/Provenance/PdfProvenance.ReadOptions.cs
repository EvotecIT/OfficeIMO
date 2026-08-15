using OfficeIMO.Provenance;

namespace OfficeIMO.Pdf;

public static partial class PdfProvenance {
    /// <summary>
    /// Combines provenance carrier limits with PDF reader limits. A caller-supplied reader limit remains
    /// authoritative; synthesized options inherit a raised provenance asset limit so both layers accept the same input.
    /// </summary>
    internal static PdfReadOptions CreateReadOptionsForInspection(
        OfficeProvenanceOptions options,
        PdfReadOptions? readOptions) {
        Guard.NotNull(options, nameof(options));
        OfficeProvenanceBinary.ValidateLimits(options);
        return CreateReadOptions(
            options.MaxAssetBytes,
            options.MaxContainerEntries,
            options.MaxExpandedContainerBytes,
            GetMaximumManifestBytes(options),
            readOptions);
    }

    private static PdfReadOptions CreateReadOptions(
        long maximumAssetBytes,
        int maximumContainerEntries,
        long maximumExpandedContainerBytes,
        long maximumManifestBytes,
        PdfReadOptions? readOptions) {
        PdfReadOptions effective = PdfReadOptions.WithMaximumContainerEntries(
            readOptions,
            maximumContainerEntries,
            maximumExpandedContainerBytes,
            maximumTotalAttachmentBytes: readOptions == null ? maximumManifestBytes : null);
        return readOptions == null
            ? PdfReadOptions.WithMinimumInputBytes(effective, maximumAssetBytes)
            : effective;
    }

    private static long GetMaximumManifestBytes(OfficeProvenanceOptions options) => Math.Min(
        options.MaxExpandedContainerBytes,
        MultiplySaturating(options.MaxManifestBytes, options.MaxCarriers));

    private static long MultiplySaturating(long value, int multiplier) =>
        value > long.MaxValue / multiplier ? long.MaxValue : value * multiplier;
}

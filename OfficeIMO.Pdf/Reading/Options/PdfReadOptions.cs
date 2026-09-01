namespace OfficeIMO.Pdf;

/// <summary>Chooses the cost and evidence depth of canonical PDF semantic reconstruction.</summary>
public enum PdfReadProfile {
    /// <summary>
    /// Uses the canonical page pipeline while omitting optional document-wide enrichment stages.
    /// The result reports every capability that was applied or skipped.
    /// </summary>
    Fast,

    /// <summary>
    /// Uses the complete built-in semantic reconstruction pipeline, including document-wide evidence.
    /// </summary>
    Structured
}

/// <summary>Controls semantic reconstruction performed by <see cref="PdfDocument.Read(PdfReadOptions, System.Threading.CancellationToken)"/>.</summary>
public sealed class PdfReadOptions {
    /// <summary>Creates independent structured-read settings.</summary>
    public static PdfReadOptions Default => new PdfReadOptions();

    /// <summary>Semantic reconstruction profile. Structured is the default public contract.</summary>
    public PdfReadProfile Profile { get; init; } = PdfReadProfile.Structured;

    /// <summary>Optional caller-ordered page selection. Null reads every page in document order.</summary>
    public PdfPageSelection? PageSelection { get; init; }

    /// <summary>Layout and geometry settings shared by every built-in semantic stage.</summary>
    public PdfTextLayoutOptions LayoutOptions { get; init; } = new PdfTextLayoutOptions();

    /// <summary>
    /// Optional semantic-stage customization. Null selects the built-in stages for <see cref="Profile"/>.
    /// Custom stages still run inside the canonical read engine and produce the same result contract.
    /// </summary>
    public PdfUnderstandingPipelineOptions? Pipeline { get; init; }

    internal static PdfReadOptions Resolve(PdfReadOptions? options) {
        PdfReadOptions effective = options ?? Default;
        if (effective.Profile < PdfReadProfile.Fast || effective.Profile > PdfReadProfile.Structured) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.Profile, "Unknown PDF read profile.");
        }
        Guard.NotNull(effective.LayoutOptions, nameof(LayoutOptions));
        return effective;
    }
}

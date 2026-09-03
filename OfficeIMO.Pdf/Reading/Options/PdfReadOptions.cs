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

    /// <summary>
    /// Creates an independent copy of these semantic read settings.
    /// Custom stage instances are reused, while mutable layout and pipeline option containers are copied.
    /// </summary>
    public PdfReadOptions Clone() {
        PdfReadOptions effective = Resolve(this);
        return new PdfReadOptions {
            Profile = effective.Profile,
            PageSelection = effective.PageSelection,
            LayoutOptions = CloneLayoutOptions(effective.LayoutOptions),
            Pipeline = ClonePipelineOptions(effective.Pipeline)
        };
    }

    internal static PdfReadOptions Resolve(PdfReadOptions? options) {
        PdfReadOptions effective = options ?? Default;
        if (effective.Profile < PdfReadProfile.Fast || effective.Profile > PdfReadProfile.Structured) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.Profile, "Unknown PDF read profile.");
        }
        Guard.NotNull(effective.LayoutOptions, nameof(LayoutOptions));
        return effective;
    }

    internal static PdfReadOptions WithPageSelection(PdfReadOptions options, PdfPageSelection? pageSelection) {
        PdfReadOptions effective = Resolve(options);
        return new PdfReadOptions {
            Profile = effective.Profile,
            PageSelection = pageSelection,
            LayoutOptions = CloneLayoutOptions(effective.LayoutOptions),
            Pipeline = ClonePipelineOptions(effective.Pipeline)
        };
    }

    private static PdfTextLayoutOptions CloneLayoutOptions(PdfTextLayoutOptions options) => new PdfTextLayoutOptions {
        MarginLeft = options.MarginLeft,
        MarginRight = options.MarginRight,
        BinWidth = options.BinWidth,
        MinGutterWidth = options.MinGutterWidth,
        LineMergeToleranceEm = options.LineMergeToleranceEm,
        LineMergeMaxPoints = options.LineMergeMaxPoints,
        ForceSingleColumn = options.ForceSingleColumn,
        ReadingDirection = options.ReadingDirection,
        JoinSoftHyphensAcrossLines = options.JoinSoftHyphensAcrossLines,
        IgnoreHeaderHeight = options.IgnoreHeaderHeight,
        IgnoreFooterHeight = options.IgnoreFooterHeight,
        GapSpaceThresholdEm = options.GapSpaceThresholdEm,
        GapGlyphFactor = options.GapGlyphFactor
    };

    private static PdfUnderstandingPipelineOptions? ClonePipelineOptions(PdfUnderstandingPipelineOptions? options) {
        if (options is null) return null;

        return new PdfUnderstandingPipelineOptions {
            GlyphDecoding = options.GlyphDecoding,
            WordGrouping = options.WordGrouping,
            LineGrouping = options.LineGrouping,
            TableDetection = options.TableDetection,
            PageSegmentation = options.PageSegmentation,
            ReadingOrder = options.ReadingOrder,
            SemanticClassification = options.SemanticClassification,
            MaxPages = options.MaxPages,
            MaxRunsPerPage = options.MaxRunsPerPage,
            MaxTextCharactersPerPage = options.MaxTextCharactersPerPage,
            MaxWordsPerPage = options.MaxWordsPerPage,
            MaxLinesPerPage = options.MaxLinesPerPage,
            MaxTableCandidatesPerPage = options.MaxTableCandidatesPerPage,
            MaxRegionsPerPage = options.MaxRegionsPerPage,
            MaxWorkUnitsPerPage = options.MaxWorkUnitsPerPage,
            MaxDocumentWorkUnits = options.MaxDocumentWorkUnits
        };
    }
}

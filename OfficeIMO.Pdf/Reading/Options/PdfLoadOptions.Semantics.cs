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

/// <summary>Controls semantic reconstruction performed by <see cref="PdfDocument.Read(PdfLoadOptions, System.Threading.CancellationToken)"/>.</summary>
public sealed partial class PdfLoadOptions {

    /// <summary>Semantic reconstruction profile. Structured is the default public contract.</summary>
    public PdfReadProfile Profile { get; init; } = PdfReadProfile.Structured;

    /// <summary>Includes image placements, image-region analysis and extracted image payloads. Default: true.
    /// False retains native text, including text inside Form XObjects, without materializing image assets.</summary>
    public bool IncludeImages { get; init; } = true;

    /// <summary>Includes logical AcroForm widget projections. Default: true.
    /// False retains bounded source field facts and values while omitting secondary widget objects.</summary>
    public bool IncludeFormWidgets { get; init; } = true;

    /// <summary>Optional caller-ordered page selection. Null reads every page in document order.</summary>
    public PdfPageSelection? PageSelection { get; set; }

    /// <summary>Layout and geometry settings shared by every built-in semantic stage.</summary>
    public PdfTextLayoutOptions LayoutOptions { get; init; } = new PdfTextLayoutOptions();

    /// <summary>
    /// Optional semantic-stage customization. Null selects the built-in stages for <see cref="Profile"/>.
    /// Custom stages still run inside the canonical read engine and produce the same result contract.
    /// </summary>
    public PdfUnderstandingPipelineOptions? Pipeline { get; init; }





    internal static PdfLoadOptions WithPageSelection(PdfLoadOptions options, PdfPageSelection? pageSelection) {
        PdfLoadOptions clone = Resolve(options).Clone();
        clone.PageSelection = pageSelection;
        return clone;
    }

    private static PdfTextLayoutOptions CloneLayoutOptions(PdfTextLayoutOptions options) => new PdfTextLayoutOptions {
        MarginLeft = options.MarginLeft,
        MarginRight = options.MarginRight,
        BinWidth = options.BinWidth,
        MinGutterWidth = options.MinGutterWidth,
        LineMergeToleranceEm = options.LineMergeToleranceEm,
        LineMergeMax = options.LineMergeMax,
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
            ImageRegionDetection = options.ImageRegionDetection,
            MaxPages = options.MaxPages,
            MaxRunsPerPage = options.MaxRunsPerPage,
            MaxTextCharactersPerPage = options.MaxTextCharactersPerPage,
            MaxWordsPerPage = options.MaxWordsPerPage,
            MaxLinesPerPage = options.MaxLinesPerPage,
            MaxTableCandidatesPerPage = options.MaxTableCandidatesPerPage,
            MaxRegionsPerPage = options.MaxRegionsPerPage,
            MaxImageRegionsPerPage = options.MaxImageRegionsPerPage,
            MaxImageCaptionCandidatesPerPage = options.MaxImageCaptionCandidatesPerPage,
            MaxWorkUnitsPerPage = options.MaxWorkUnitsPerPage,
            MaxDocumentWorkUnits = options.MaxDocumentWorkUnits
        };
    }
}

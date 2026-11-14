namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a parsed PDF document with access to pages, catalog and metadata.
/// Note: MVP reader supports classic xref tables and simple stream parsing sufficient for OfficeIMO.Pdf output.
/// </summary>
internal sealed partial class PdfReadDocument {
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly PdfDictionary? _catalog;
    private readonly string _trailerRaw;
    private readonly PdfLoadOptions _options;
    private readonly long _decodedStreamBytes;
    private readonly PdfDecodedStreamBudget _decodedStreamBudget;
    private Dictionary<int, int>? _pageNumberByObject;
    private readonly PdfFontResourceCache _fontResourceCache = new();
    private readonly PdfDocumentProperties _metadata;
    private readonly IReadOnlyList<PdfOutputIntentInfo> _outputIntents;
    private readonly bool _outputIntentsAreComplete;
    private readonly PdfIccProfileRetentionBudget _outputIntentMetadataRetentionBudget;
    private readonly IReadOnlyList<PdfPageLabel> _pageLabels;
    private readonly IReadOnlyList<PdfCatalogAction> _catalogActions;
    private readonly IReadOnlyList<PdfJavaScript> _javaScripts;
    private readonly PdfOptionalContentProperties? _optionalContent;
    private readonly PdfDocumentOpenAction? _openAction;
    private readonly string? _acroFormDefaultAppearance;
    private readonly int? _acroFormQuadding;
    private readonly bool? _acroFormNeedAppearances;
    private readonly int? _acroFormSignatureFlags;
    private readonly PdfOutputIntentColorTransform? _outputIntentColorTransform;
    private readonly PdfPageOptionalContentVisibility.DocumentState _optionalContentVisibilityState;

    internal Dictionary<int, PdfIndirectObject> Objects => _objects;
    internal string TrailerRaw => _trailerRaw;
    internal PdfDictionary? CatalogDictionary => FindCatalog();
    internal PdfLoadOptions ReadOptions => _options;
    internal int FormWidgetJavaScriptCount => _formWidgetJavaScriptCount;
    internal long FormWidgetJavaScriptBytes => _formWidgetJavaScriptBytes;
    internal long DecodedStreamBytes => _decodedStreamBytes;
    internal PdfDecodedStreamBudget DecodedStreamBudget => _decodedStreamBudget;

    private PdfReadDocument(
        Dictionary<int, PdfIndirectObject> objects,
        string trailerRaw,
        PdfDocumentSecurityInfo security,
        PdfRepairReport repairReport,
        PdfLoadOptions? options,
        long decodedStreamBytes,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (objects.Count == 0) {
            throw new PdfParseException("no-indirect-objects", "The PDF contains no readable indirect objects and cannot be opened.", null);
        }
        _objects = objects; _trailerRaw = trailerRaw; _options = options ?? new PdfLoadOptions();
        _namedDestinationsFeature = Defer(ExtractNamedDestinations);
        _attachmentsFeature = Defer(ExtractAttachmentInfos);
        _xmpMetadataFeature = Defer(ExtractXmpMetadata);
        _taggedContentFeature = Defer(ExtractTaggedContent);
        _outlinesFeature = Defer(ExtractOutlines);
        _portfolioFeature = Defer(ExtractPortfolio);
        _formsFeature = Defer(ExtractFormState);
        _decodedStreamBudget = new PdfDecodedStreamBudget(_options.Limits, decodedStreamBytes);
        _outputIntentMetadataRetentionBudget = new PdfIccProfileRetentionBudget(_options.Limits.MaxDecodedStreamBytes);
        Security = security;
        _catalog = PdfSyntax.FindCatalog(_objects, _trailerRaw, cancellationToken);
        _optionalContentVisibilityState = PdfPageOptionalContentVisibility.CreateDocumentState(
            _catalog,
            _objects,
            _options.Limits.MaxContentNestingDepth,
            cancellationToken);
        _outputIntentColorTransform = PdfOutputIntentColorTransform.TryCreate(
            _catalog,
            _objects,
            _options.Limits.MaxDecodedStreamBytes,
            _outputIntentMetadataRetentionBudget,
            cancellationToken);
        Pages = CollectPages(cancellationToken);
        PdfRepairReport loadRepairs = repairReport.Append(PdfSemanticRepairDiagnostics.AnalyzeAndRepair(_objects, _catalog, Pages, _options, cancellationToken));
        _repairReportFeature = Defer(token => loadRepairs.Append(
            PdfSemanticRepairDiagnostics.AnalyzeOrphanedSemanticObjects(_objects, _catalog, token)));
        cancellationToken.ThrowIfCancellationRequested();
        _metadata = ExtractMetadata(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        _pageLabels = ExtractPageLabels(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        _catalogActions = ExtractCatalogActions(out _javaScripts, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        _outputIntents = ExtractOutputIntents(out bool outputIntentsAreComplete, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        _outputIntentsAreComplete = outputIntentsAreComplete;
        _optionalContent = ExtractOptionalContent(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        _openAction = ExtractOpenAction(cancellationToken);
        ViewerPreferences = ExtractViewerPreferences(cancellationToken);
        _acroFormDefaultAppearance = ExtractAcroFormText("DA");
        _acroFormQuadding = ExtractAcroFormInteger("Q");
        _acroFormNeedAppearances = ExtractAcroFormBoolean("NeedAppearances");
        _acroFormSignatureFlags = ExtractAcroFormInteger("SigFlags");
        CatalogPageMode = ExtractCatalogName("PageMode");
        CatalogPageLayout = ExtractCatalogName("PageLayout");
        CatalogVersion = ExtractCatalogName("Version");
        CatalogLanguage = ExtractCatalogString("Lang");
        if (_options.ParsingMode == PdfParsingMode.Strict) {
            CompleteDeferredFeatures(cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
        _decodedStreamBytes = _decodedStreamBudget.UsedBytes;
    }

    /// <summary>All page objects discovered in document order.</summary>
    public IReadOnlyList<PdfReadPage> Pages { get; private set; }

    internal PdfReadDocument WithInspectionPages(PdfReadPage[] pages) {
        // The inspection owns its page array. Parsed objects and source inventory remain shared.
        var inspection = (PdfReadDocument)MemberwiseClone();
        inspection.Pages = Array.AsReadOnly(pages);
        return inspection;
    }

    /// <summary>Document metadata (when present).</summary>
    public PdfDocumentProperties Metadata => ReadLogicalContent(_metadata);

    /// <summary>Top-level document outline/bookmark entries.</summary>
    public IReadOnlyList<PdfOutlineItem> Outlines => ReadLogicalFeature(_outlinesFeature);

    /// <summary>Page-label rules discovered from the document catalog.</summary>
    public IReadOnlyList<PdfPageLabel> PageLabels => ReadLogicalContent(_pageLabels);

    /// <summary>Named destinations discovered from the document catalog.</summary>
    public IReadOnlyList<PdfNamedDestination> NamedDestinations => ReadLogicalFeature(_namedDestinationsFeature).Destinations;

    /// <summary>Catalog-level actions discovered from supported name trees.</summary>
    public IReadOnlyList<PdfCatalogAction> CatalogActions => ReadLogicalContent(_catalogActions);

    /// <summary>Named document-level JavaScript actions discovered from the catalog name tree.</summary>
    public IReadOnlyList<PdfJavaScript> JavaScripts => ReadLogicalContent(_javaScripts);

    /// <summary>Simple document open action discovered from the document catalog, when supported.</summary>
    public PdfDocumentOpenAction? OpenAction => ReadLogicalContent(_openAction);

    /// <summary>Simple viewer preference entries discovered from the document catalog, when supported.</summary>
    public PdfViewerPreferences? ViewerPreferences { get; }

    /// <summary>Document portfolio metadata discovered from the catalog, when present.</summary>
    public PdfPortfolioInfo? Portfolio => ReadLogicalFeature(_portfolioFeature);

    /// <summary>Simple AcroForm fields discovered from the document catalog.</summary>
    public IReadOnlyList<PdfFormField> FormFields => ReadLogicalFeature(_formsFeature).Fields;

    /// <summary>AcroForm default appearance string from /DA, when present.</summary>
    public string? AcroFormDefaultAppearance => ReadLogicalContent(_acroFormDefaultAppearance);

    /// <summary>Raw AcroForm default /Q quadding value, when present.</summary>
    public int? AcroFormQuadding => ReadLogicalContent(_acroFormQuadding);

    /// <summary>AcroForm NeedAppearances flag, when present.</summary>
    public bool? AcroFormNeedAppearances => ReadLogicalContent(_acroFormNeedAppearances);

    /// <summary>Raw AcroForm signature flags from /SigFlags, when present.</summary>
    public int? AcroFormSignatureFlags => ReadLogicalContent(_acroFormSignatureFlags);

    /// <summary>AcroForm XFA packet metadata when the document catalog exposes /AcroForm /XFA.</summary>
    public PdfAcroFormXfaInfo? AcroFormXfa => ReadLogicalFeature(_formsFeature).Xfa;

    /// <summary>Catalog page mode, for example UseOutlines or FullScreen, when present.</summary>
    public string? CatalogPageMode { get; }

    /// <summary>Catalog page layout, for example SinglePage or TwoColumnLeft, when present.</summary>
    public string? CatalogPageLayout { get; }

    /// <summary>Catalog PDF version override, for example 1.7, when present.</summary>
    public string? CatalogVersion { get; }

    /// <summary>Catalog language tag, for example en-US or pl-PL, when present.</summary>
    public string? CatalogLanguage { get; }

    /// <summary>Security, signature, and revision markers read from the source PDF bytes.</summary>
    public PdfDocumentSecurityInfo Security { get; }

    /// <summary>Structural recoveries applied while loading this document.</summary>
    public PdfRepairReport RepairReport => _repairReportFeature.GetValue();

    internal IReadOnlyList<PdfOutlineItem> UncheckedOutlines => _outlines;
    // Rewrites and inspection must read the parsed source, since Metadata is a mutable read model.
    internal PdfDocumentProperties UncheckedMetadata => ExtractMetadata();
    internal PdfDocumentProperties GetUncheckedMetadata(CancellationToken cancellationToken) => ExtractMetadata(cancellationToken);
    internal PdfXmpMetadataInfo? UncheckedXmpMetadata => _xmpMetadata;
    internal IReadOnlyList<PdfOutputIntentInfo> UncheckedOutputIntents => _outputIntents;
    internal bool UncheckedOutputIntentsAreComplete => _outputIntentsAreComplete;
    internal IReadOnlyList<PdfPageLabel> UncheckedPageLabels => _pageLabels;
    internal IReadOnlyList<PdfNamedDestination> UncheckedNamedDestinations => _namedDestinations;
    internal IReadOnlyList<PdfCatalogAction> UncheckedCatalogActions => _catalogActions;
    internal IReadOnlyList<PdfJavaScript> UncheckedJavaScripts => _javaScripts;
    internal IReadOnlyList<PdfAttachmentInfo> UncheckedAttachments => _attachments;
    internal PdfTaggedContentInfo? UncheckedTaggedContent => _taggedContent;
    internal PdfOptionalContentProperties? UncheckedOptionalContent => _optionalContent;
    internal PdfDocumentOpenAction? UncheckedOpenAction => _openAction;
    internal IReadOnlyList<PdfFormField> UncheckedFormFields => _formFields;
    internal string? UncheckedAcroFormDefaultAppearance => _acroFormDefaultAppearance;
    internal int? UncheckedAcroFormQuadding => _acroFormQuadding;
    internal bool? UncheckedAcroFormNeedAppearances => _acroFormNeedAppearances;
    internal int? UncheckedAcroFormSignatureFlags => _acroFormSignatureFlags;
    internal PdfAcroFormXfaInfo? UncheckedAcroFormXfa => _acroFormXfa;

    private T ReadLogicalContent<T>(T value) {
        DemandContentExtraction("logical object");
        return value;
    }

    private T ReadLogicalFeature<T>(DocumentFeature<T> feature) {
        DemandContentExtraction("logical object");
        return feature.GetValue();
    }
}

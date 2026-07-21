namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a parsed PDF document with access to pages, catalog and metadata.
/// Note: MVP reader supports classic xref tables and simple stream parsing sufficient for OfficeIMO.Pdf output.
/// </summary>
public sealed partial class PdfReadDocument {
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly string _trailerRaw;
    private readonly PdfReadOptions _options;
    private readonly Dictionary<string, PdfNamedDestination> _nameDestinations = new(StringComparer.Ordinal);
    private readonly Dictionary<string, PdfNamedDestination> _stringDestinations = new(StringComparer.Ordinal);
    private readonly PdfMetadata _metadata;
    private readonly PdfXmpMetadataInfo? _xmpMetadata;
    private readonly IReadOnlyList<PdfOutputIntentInfo> _outputIntents;
    private readonly IReadOnlyList<PdfOutlineItem> _outlines;
    private readonly IReadOnlyList<PdfPageLabel> _pageLabels;
    private readonly IReadOnlyList<PdfNamedDestination> _namedDestinations;
    private readonly IReadOnlyList<PdfCatalogAction> _catalogActions;
    private readonly IReadOnlyList<PdfAttachmentInfo> _attachments;
    private readonly PdfTaggedContentInfo? _taggedContent;
    private readonly PdfOptionalContentProperties? _optionalContent;
    private readonly PdfDocumentOpenAction? _openAction;
    private readonly PdfPortfolioInfo? _portfolio;
    private readonly IReadOnlyList<PdfFormField> _formFields;
    private readonly string? _acroFormDefaultAppearance;
    private readonly int? _acroFormQuadding;
    private readonly bool? _acroFormNeedAppearances;
    private readonly int? _acroFormSignatureFlags;
    private readonly PdfAcroFormXfaInfo? _acroFormXfa;

    internal Dictionary<int, PdfIndirectObject> Objects => _objects;
    internal string TrailerRaw => _trailerRaw;
    internal PdfReadOptions ReadOptions => _options;

    private PdfReadDocument(
        Dictionary<int, PdfIndirectObject> objects,
        string trailerRaw,
        PdfDocumentSecurityInfo security,
        PdfRepairReport repairReport,
        PdfReadOptions? options) {
        _objects = objects; _trailerRaw = trailerRaw; _options = options ?? new PdfReadOptions();
        Security = security;
        Pages = CollectPages();
        RepairReport = repairReport.Append(PdfSemanticRepairDiagnostics.AnalyzeAndRepair(_objects, FindCatalog(), Pages, _options));
        _metadata = ExtractMetadata();
        _pageLabels = ExtractPageLabels();
        _namedDestinations = ExtractNamedDestinations();
        _catalogActions = ExtractCatalogActions();
        _attachments = ExtractAttachmentInfos();
        _outputIntents = ExtractOutputIntents();
        _xmpMetadata = ExtractXmpMetadata();
        _taggedContent = ExtractTaggedContent();
        _optionalContent = ExtractOptionalContent();
        _outlines = ExtractOutlines();
        _openAction = ExtractOpenAction();
        ViewerPreferences = ExtractViewerPreferences();
        _portfolio = ExtractPortfolio();
        _acroFormDefaultAppearance = ExtractAcroFormText("DA");
        _acroFormQuadding = ExtractAcroFormInteger("Q");
        _acroFormXfa = ExtractAcroFormXfaInfo();
        _formFields = ExtractFormFields();
        _acroFormNeedAppearances = ExtractAcroFormBoolean("NeedAppearances");
        _acroFormSignatureFlags = ExtractAcroFormInteger("SigFlags");
        CatalogPageMode = ExtractCatalogName("PageMode");
        CatalogPageLayout = ExtractCatalogName("PageLayout");
        CatalogVersion = ExtractCatalogName("Version");
        CatalogLanguage = ExtractCatalogString("Lang");
    }

    /// <summary>All page objects discovered in document order.</summary>
    public IReadOnlyList<PdfReadPage> Pages { get; }

    /// <summary>Document metadata (when present).</summary>
    public PdfMetadata Metadata => ReadLogicalContent(_metadata);

    /// <summary>Top-level document outline/bookmark entries.</summary>
    public IReadOnlyList<PdfOutlineItem> Outlines => ReadLogicalContent(_outlines);

    /// <summary>Page-label rules discovered from the document catalog.</summary>
    public IReadOnlyList<PdfPageLabel> PageLabels => ReadLogicalContent(_pageLabels);

    /// <summary>Named destinations discovered from the document catalog.</summary>
    public IReadOnlyList<PdfNamedDestination> NamedDestinations => ReadLogicalContent(_namedDestinations);

    /// <summary>Catalog-level actions discovered from supported name trees.</summary>
    public IReadOnlyList<PdfCatalogAction> CatalogActions => ReadLogicalContent(_catalogActions);

    /// <summary>Simple document open action discovered from the document catalog, when supported.</summary>
    public PdfDocumentOpenAction? OpenAction => ReadLogicalContent(_openAction);

    /// <summary>Simple viewer preference entries discovered from the document catalog, when supported.</summary>
    public PdfViewerPreferences? ViewerPreferences { get; }

    /// <summary>Document portfolio metadata discovered from the catalog, when present.</summary>
    public PdfPortfolioInfo? Portfolio => ReadLogicalContent(_portfolio);

    /// <summary>Simple AcroForm fields discovered from the document catalog.</summary>
    public IReadOnlyList<PdfFormField> FormFields => ReadLogicalContent(_formFields);

    /// <summary>AcroForm default appearance string from /DA, when present.</summary>
    public string? AcroFormDefaultAppearance => ReadLogicalContent(_acroFormDefaultAppearance);

    /// <summary>Raw AcroForm default /Q quadding value, when present.</summary>
    public int? AcroFormQuadding => ReadLogicalContent(_acroFormQuadding);

    /// <summary>AcroForm NeedAppearances flag, when present.</summary>
    public bool? AcroFormNeedAppearances => ReadLogicalContent(_acroFormNeedAppearances);

    /// <summary>Raw AcroForm signature flags from /SigFlags, when present.</summary>
    public int? AcroFormSignatureFlags => ReadLogicalContent(_acroFormSignatureFlags);

    /// <summary>AcroForm XFA packet metadata when the document catalog exposes /AcroForm /XFA.</summary>
    public PdfAcroFormXfaInfo? AcroFormXfa => ReadLogicalContent(_acroFormXfa);

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
    public PdfRepairReport RepairReport { get; }

    internal IReadOnlyList<PdfOutlineItem> UncheckedOutlines => _outlines;
    internal PdfMetadata UncheckedMetadata => _metadata;
    internal PdfXmpMetadataInfo? UncheckedXmpMetadata => _xmpMetadata;
    internal IReadOnlyList<PdfOutputIntentInfo> UncheckedOutputIntents => _outputIntents;
    internal IReadOnlyList<PdfPageLabel> UncheckedPageLabels => _pageLabels;
    internal IReadOnlyList<PdfNamedDestination> UncheckedNamedDestinations => _namedDestinations;
    internal IReadOnlyList<PdfCatalogAction> UncheckedCatalogActions => _catalogActions;
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
}

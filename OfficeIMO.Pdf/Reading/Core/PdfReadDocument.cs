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
        Metadata = ExtractMetadata();
        PageLabels = ExtractPageLabels();
        NamedDestinations = ExtractNamedDestinations();
        CatalogActions = ExtractCatalogActions();
        Attachments = ExtractAttachmentInfos();
        OutputIntents = ExtractOutputIntents();
        XmpMetadata = ExtractXmpMetadata();
        TaggedContent = ExtractTaggedContent();
        OptionalContent = ExtractOptionalContent();
        Outlines = ExtractOutlines();
        OpenAction = ExtractOpenAction();
        ViewerPreferences = ExtractViewerPreferences();
        Portfolio = ExtractPortfolio();
        AcroFormDefaultAppearance = ExtractAcroFormText("DA");
        AcroFormQuadding = ExtractAcroFormInteger("Q");
        AcroFormXfa = ExtractAcroFormXfaInfo();
        FormFields = ExtractFormFields();
        AcroFormNeedAppearances = ExtractAcroFormBoolean("NeedAppearances");
        AcroFormSignatureFlags = ExtractAcroFormInteger("SigFlags");
        CatalogPageMode = ExtractCatalogName("PageMode");
        CatalogPageLayout = ExtractCatalogName("PageLayout");
        CatalogVersion = ExtractCatalogName("Version");
        CatalogLanguage = ExtractCatalogString("Lang");
    }

    /// <summary>All page objects discovered in document order.</summary>
    public IReadOnlyList<PdfReadPage> Pages { get; }

    /// <summary>Document metadata (when present).</summary>
    public PdfMetadata Metadata { get; }

    /// <summary>Top-level document outline/bookmark entries.</summary>
    public IReadOnlyList<PdfOutlineItem> Outlines { get; }

    /// <summary>Page-label rules discovered from the document catalog.</summary>
    public IReadOnlyList<PdfPageLabel> PageLabels { get; }

    /// <summary>Named destinations discovered from the document catalog.</summary>
    public IReadOnlyList<PdfNamedDestination> NamedDestinations { get; }

    /// <summary>Catalog-level actions discovered from supported name trees.</summary>
    public IReadOnlyList<PdfCatalogAction> CatalogActions { get; }

    /// <summary>Simple document open action discovered from the document catalog, when supported.</summary>
    public PdfDocumentOpenAction? OpenAction { get; }

    /// <summary>Simple viewer preference entries discovered from the document catalog, when supported.</summary>
    public PdfViewerPreferences? ViewerPreferences { get; }

    /// <summary>Document portfolio metadata discovered from the catalog, when present.</summary>
    public PdfPortfolioInfo? Portfolio { get; }

    /// <summary>Simple AcroForm fields discovered from the document catalog.</summary>
    public IReadOnlyList<PdfFormField> FormFields { get; }

    /// <summary>AcroForm default appearance string from /DA, when present.</summary>
    public string? AcroFormDefaultAppearance { get; }

    /// <summary>Raw AcroForm default /Q quadding value, when present.</summary>
    public int? AcroFormQuadding { get; }

    /// <summary>AcroForm NeedAppearances flag, when present.</summary>
    public bool? AcroFormNeedAppearances { get; }

    /// <summary>Raw AcroForm signature flags from /SigFlags, when present.</summary>
    public int? AcroFormSignatureFlags { get; }

    /// <summary>AcroForm XFA packet metadata when the document catalog exposes /AcroForm /XFA.</summary>
    public PdfAcroFormXfaInfo? AcroFormXfa { get; }

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
}

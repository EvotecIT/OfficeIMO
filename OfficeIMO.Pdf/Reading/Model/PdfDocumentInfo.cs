namespace OfficeIMO.Pdf;

/// <summary>
/// Basic document-level information useful for inspection and automation scenarios.
/// </summary>
public sealed class PdfDocumentInfo {
    private IReadOnlyList<PdfLinkAnnotation>? _linkAnnotations;
    private IReadOnlyList<string>? _linkUris;
    private IReadOnlyList<string>? _linkDestinationNames;
    private IReadOnlyList<string>? _namedDestinationNames;
    private IReadOnlyList<string>? _formFieldNames;

    internal PdfDocumentInfo(IReadOnlyList<PdfPageInfo> pages, PdfMetadata metadata, IReadOnlyList<PdfOutlineItem> outlines, IReadOnlyList<PdfPageLabel> pageLabels, IReadOnlyList<PdfNamedDestination> namedDestinations, PdfDocumentOpenAction? openAction, PdfViewerPreferences? viewerPreferences, IReadOnlyList<PdfFormField> formFields, string? acroFormDefaultAppearance, bool? acroFormNeedAppearances, int? acroFormSignatureFlags, string? headerVersion, string? catalogPageMode, string? catalogPageLayout, string? catalogVersion, string? catalogLanguage, bool hasSignatures, bool hasForms, bool hasAnnotations, bool hasOutlines, bool hasCatalogViewSettings, bool hasPageLabels, bool hasCatalogNameTrees, bool hasNamedDestinations, bool hasOpenActions, bool hasViewerPreferences, bool hasTaggedContent, bool hasXmpMetadata, bool hasCatalogUri, bool hasOutputIntents, bool hasEmbeddedFiles, bool hasOptionalContent, bool hasActiveContent) {
        Pages = pages;
        Metadata = metadata;
        Outlines = outlines;
        PageLabels = pageLabels;
        NamedDestinations = namedDestinations;
        OpenAction = openAction;
        ViewerPreferences = viewerPreferences;
        FormFields = formFields;
        AcroFormDefaultAppearance = acroFormDefaultAppearance;
        AcroFormNeedAppearances = acroFormNeedAppearances;
        AcroFormSignatureFlags = acroFormSignatureFlags;
        HeaderVersion = headerVersion;
        CatalogPageMode = catalogPageMode;
        CatalogPageLayout = catalogPageLayout;
        CatalogVersion = catalogVersion;
        CatalogLanguage = catalogLanguage;
        HasSignatures = hasSignatures;
        HasForms = hasForms;
        HasAnnotations = hasAnnotations;
        HasOutlines = hasOutlines;
        HasCatalogViewSettings = hasCatalogViewSettings;
        HasPageLabels = hasPageLabels;
        HasCatalogNameTrees = hasCatalogNameTrees;
        HasNamedDestinations = hasNamedDestinations;
        HasOpenActions = hasOpenActions;
        HasViewerPreferences = hasViewerPreferences;
        HasTaggedContent = hasTaggedContent;
        HasXmpMetadata = hasXmpMetadata;
        HasCatalogUri = hasCatalogUri;
        HasOutputIntents = hasOutputIntents;
        HasEmbeddedFiles = hasEmbeddedFiles;
        HasOptionalContent = hasOptionalContent;
        HasActiveContent = hasActiveContent;
    }

    /// <summary>Number of pages in the document.</summary>
    public int PageCount => Pages.Count;

    /// <summary>Pages in document order.</summary>
    public IReadOnlyList<PdfPageInfo> Pages { get; }

    /// <summary>Number of simple link annotations read from all pages.</summary>
    public int LinkAnnotationCount => LinkAnnotations.Count;

    /// <summary>Number of distinct simple URI link targets read from all pages.</summary>
    public int LinkUriCount => LinkUris.Count;

    /// <summary>Number of distinct simple named-destination link targets read from all pages.</summary>
    public int LinkDestinationCount => LinkDestinationNames.Count;

    /// <summary>Number of named destinations read from the document catalog.</summary>
    public int NamedDestinationCount => NamedDestinations.Count;

    /// <summary>Number of simple AcroForm fields read from the document catalog.</summary>
    public int FormFieldCount => FormFields.Count;

    /// <summary>Number of page-label rules read from the document catalog.</summary>
    public int PageLabelCount => PageLabels.Count;

    /// <summary>Simple link annotations read from all pages in document order.</summary>
    public IReadOnlyList<PdfLinkAnnotation> LinkAnnotations {
        get {
            if (_linkAnnotations is not null) {
                return _linkAnnotations;
            }

            var links = new List<PdfLinkAnnotation>();
            for (int i = 0; i < Pages.Count; i++) {
                for (int j = 0; j < Pages[i].LinkAnnotations.Count; j++) {
                    var link = Pages[i].LinkAnnotations[j];
                    links.Add(link.PageNumber.HasValue ? link : link.WithPageNumber(Pages[i].PageNumber));
                }
            }

            _linkAnnotations = links.AsReadOnly();
            return _linkAnnotations;
        }
    }

    /// <summary>Distinct simple URI link targets read from all pages in first-seen document order.</summary>
    public IReadOnlyList<string> LinkUris {
        get {
            if (_linkUris is not null) {
                return _linkUris;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var uris = new List<string>();
            foreach (var link in LinkAnnotations) {
                if (link.Uri != null && seen.Add(link.Uri)) {
                    uris.Add(link.Uri);
                }
            }

            _linkUris = uris.AsReadOnly();
            return _linkUris;
        }
    }

    /// <summary>Distinct simple named-destination link targets read from all pages in first-seen document order.</summary>
    public IReadOnlyList<string> LinkDestinationNames {
        get {
            if (_linkDestinationNames is not null) {
                return _linkDestinationNames;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var names = new List<string>();
            foreach (var link in LinkAnnotations) {
                if (link.DestinationName != null && seen.Add(link.DestinationName)) {
                    names.Add(link.DestinationName);
                }
            }

            _linkDestinationNames = names.AsReadOnly();
            return _linkDestinationNames;
        }
    }

    /// <summary>Named destination names read from the document catalog in first-seen order.</summary>
    public IReadOnlyList<string> NamedDestinationNames {
        get {
            if (_namedDestinationNames is not null) {
                return _namedDestinationNames;
            }

            var names = new List<string>(NamedDestinations.Count);
            for (int i = 0; i < NamedDestinations.Count; i++) {
                names.Add(NamedDestinations[i].Name);
            }

            _namedDestinationNames = names.AsReadOnly();
            return _namedDestinationNames;
        }
    }

    /// <summary>Readable AcroForm field names in first-seen document order.</summary>
    public IReadOnlyList<string> FormFieldNames {
        get {
            if (_formFieldNames is not null) {
                return _formFieldNames;
            }

            var names = new List<string>();
            for (int i = 0; i < FormFields.Count; i++) {
                if (!string.IsNullOrEmpty(FormFields[i].Name)) {
                    names.Add(FormFields[i].Name!);
                }
            }

            _formFieldNames = names.AsReadOnly();
            return _formFieldNames;
        }
    }

    /// <summary>True when at least one simple link annotation was read from the document pages.</summary>
    public bool HasLinkAnnotations => LinkAnnotationCount > 0;

    /// <summary>Document metadata from the PDF Info dictionary when available.</summary>
    public PdfMetadata Metadata { get; }

    /// <summary>Top-level document outline/bookmark entries.</summary>
    public IReadOnlyList<PdfOutlineItem> Outlines { get; }

    /// <summary>Page-label rules read from the document catalog.</summary>
    public IReadOnlyList<PdfPageLabel> PageLabels { get; }

    /// <summary>True when simple page-label rules were read from the document catalog.</summary>
    public bool HasReadablePageLabels => PageLabelCount > 0;

    /// <summary>Named destinations read from the document catalog.</summary>
    public IReadOnlyList<PdfNamedDestination> NamedDestinations { get; }

    /// <summary>Simple AcroForm fields read from the document catalog.</summary>
    public IReadOnlyList<PdfFormField> FormFields { get; }

    /// <summary>True when at least one simple AcroForm field was read from the document catalog.</summary>
    public bool HasReadableFormFields => FormFieldCount > 0;

    /// <summary>AcroForm default appearance string from /DA, when present.</summary>
    public string? AcroFormDefaultAppearance { get; }

    /// <summary>True when an AcroForm default appearance string was readable.</summary>
    public bool HasAcroFormDefaultAppearance => !string.IsNullOrEmpty(AcroFormDefaultAppearance);

    /// <summary>AcroForm NeedAppearances flag, when present.</summary>
    public bool? AcroFormNeedAppearances { get; }

    /// <summary>True when the AcroForm requests viewer-side appearance regeneration.</summary>
    public bool RequiresAcroFormAppearanceRegeneration => AcroFormNeedAppearances == true;

    /// <summary>True when an AcroForm NeedAppearances flag was readable.</summary>
    public bool HasAcroFormNeedAppearances => AcroFormNeedAppearances.HasValue;

    /// <summary>Raw AcroForm signature flags from /SigFlags, when present.</summary>
    public int? AcroFormSignatureFlags { get; }

    /// <summary>True when AcroForm signature flags were readable.</summary>
    public bool HasAcroFormSignatureFlags => AcroFormSignatureFlags.HasValue;

    /// <summary>Simple document open action read from the document catalog, when supported.</summary>
    public PdfDocumentOpenAction? OpenAction { get; }

    /// <summary>True when a simple document open action was read from the document catalog.</summary>
    public bool HasReadableOpenAction => OpenAction is not null;

    /// <summary>Simple viewer preference entries read from the document catalog, when supported.</summary>
    public PdfViewerPreferences? ViewerPreferences { get; }

    /// <summary>True when simple viewer preference entries were read from the document catalog.</summary>
    public bool HasReadableViewerPreferences => ViewerPreferences is not null;

    /// <summary>PDF header version, for example 1.4, when a header is present.</summary>
    public string? HeaderVersion { get; }

    /// <summary>Catalog page mode, for example UseOutlines or FullScreen, when present.</summary>
    public string? CatalogPageMode { get; }

    /// <summary>Catalog page layout, for example SinglePage or TwoColumnLeft, when present.</summary>
    public string? CatalogPageLayout { get; }

    /// <summary>Catalog PDF version override, for example 1.7, when present.</summary>
    public string? CatalogVersion { get; }

    /// <summary>Catalog language tag, for example en-US or pl-PL, when present.</summary>
    public string? CatalogLanguage { get; }

    /// <summary>True when the document contains digital signature markers.</summary>
    public bool HasSignatures { get; }

    /// <summary>True when the document contains AcroForm or form-field markers.</summary>
    public bool HasForms { get; }

    /// <summary>True when the document contains annotation markers.</summary>
    public bool HasAnnotations { get; }

    /// <summary>True when the document contains outline/bookmark markers.</summary>
    public bool HasOutlines { get; }

    /// <summary>True when the document contains catalog page mode or layout markers.</summary>
    public bool HasCatalogViewSettings { get; }

    /// <summary>True when the document contains page label markers.</summary>
    public bool HasPageLabels { get; }

    /// <summary>True when the document contains catalog name-tree markers.</summary>
    public bool HasCatalogNameTrees { get; }

    /// <summary>True when the document contains named destination markers.</summary>
    public bool HasNamedDestinations { get; }

    /// <summary>True when the document contains document open action markers.</summary>
    public bool HasOpenActions { get; }

    /// <summary>True when the document contains viewer preference markers.</summary>
    public bool HasViewerPreferences { get; }

    /// <summary>True when the document contains tagged PDF structure markers.</summary>
    public bool HasTaggedContent { get; }

    /// <summary>True when the document contains XMP metadata stream markers.</summary>
    public bool HasXmpMetadata { get; }

    /// <summary>True when the document catalog contains a URI dictionary.</summary>
    public bool HasCatalogUri { get; }

    /// <summary>True when the document contains output intent markers.</summary>
    public bool HasOutputIntents { get; }

    /// <summary>True when the document contains embedded file markers.</summary>
    public bool HasEmbeddedFiles { get; }

    /// <summary>True when the document contains optional content/layer markers.</summary>
    public bool HasOptionalContent { get; }

    /// <summary>True when the document contains active content markers such as JavaScript actions.</summary>
    public bool HasActiveContent { get; }
}

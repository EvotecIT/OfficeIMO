namespace OfficeIMO.Pdf;

/// <summary>
/// Logical PDF element categories exposed by the first-party read model.
/// </summary>
public enum PdfLogicalElementKind {
    /// <summary>Line-level text recovered from positioned PDF text spans.</summary>
    TextBlock,
    /// <summary>Heuristic heading line inferred from text size and geometry.</summary>
    Heading,
    /// <summary>Detected bullet or numbered list item.</summary>
    ListItem,
    /// <summary>Detected leader row such as label plus dotted value.</summary>
    LeaderRow,
    /// <summary>Detected table-like region.</summary>
    Table,
    /// <summary>Image XObject referenced by the page.</summary>
    Image,
    /// <summary>URI, named-destination, direct-destination, named-action, or remote GoTo link annotation on the page.</summary>
    LinkAnnotation,
    /// <summary>AcroForm widget annotation on the page.</summary>
    FormWidget
}

/// <summary>
/// Common shape for logical page elements extracted from a PDF page.
/// </summary>
public interface IPdfLogicalElement {
    /// <summary>One-based source page number.</summary>
    int PageNumber { get; }

    /// <summary>Element kind.</summary>
    PdfLogicalElementKind Kind { get; }
}

/// <summary>
/// First-party logical read model for a parser-supported PDF.
/// </summary>
public sealed partial class PdfLogicalDocument {
    private const int AcroFormSignaturesExistFlag = 1;
    private const int AcroFormAppendOnlyFlag = 2;
    private IReadOnlyDictionary<int, IReadOnlyList<PdfLogicalPage>>? _pagesBySourcePageNumber;
    private IReadOnlyList<IPdfLogicalElement>? _elements;
    private IReadOnlyDictionary<PdfLogicalElementKind, IReadOnlyList<IPdfLogicalElement>>? _elementsByKind;
    private IReadOnlyDictionary<int, IReadOnlyList<IPdfLogicalElement>>? _elementsByPageNumber;
    private IReadOnlyList<PdfLogicalTextBlock>? _textBlocks;
    private IReadOnlyList<PdfLogicalHeading>? _headings;
    private IReadOnlyList<PdfLogicalParagraph>? _paragraphs;
    private IReadOnlyList<PdfLogicalListItem>? _listItems;
    private IReadOnlyList<PdfLogicalTable>? _tables;
    private IReadOnlyList<PdfLogicalImage>? _images;
    private IReadOnlyList<PdfLogicalLinkAnnotation>? _links;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfLogicalLinkAnnotation>>? _linksByUri;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfLogicalLinkAnnotation>>? _linksByDestinationName;
    private IReadOnlyDictionary<int, IReadOnlyList<PdfLogicalLinkAnnotation>>? _linksByDestinationPageNumber;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfLogicalLinkAnnotation>>? _linksByNamedAction;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfLogicalLinkAnnotation>>? _linksByRemoteFile;
    private IReadOnlyList<PdfLogicalFormWidget>? _formWidgets;
    private IReadOnlyDictionary<string, PdfFormField>? _formFieldsByName;
    private IReadOnlyDictionary<PdfFormFieldKind, IReadOnlyList<PdfFormField>>? _formFieldsByKind;
    private IReadOnlyList<string>? _formFieldNames;
    private IReadOnlyDictionary<int, IReadOnlyList<PdfFormField>>? _formFieldsByPageNumber;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfLogicalFormWidget>>? _formWidgetsByFieldName;
    private IReadOnlyDictionary<int, IReadOnlyList<PdfLogicalFormWidget>>? _formWidgetsByPageNumber;

    private PdfLogicalDocument(
        PdfMetadata metadata,
        IReadOnlyList<PdfLogicalPage> pages,
        IReadOnlyList<PdfOutlineItem> outlines,
        IReadOnlyList<PdfPageLabel> pageLabels,
        IReadOnlyList<PdfNamedDestination> namedDestinations,
        IReadOnlyList<PdfCatalogAction> catalogActions,
        IReadOnlyList<PdfAttachmentInfo> attachments,
        IReadOnlyList<PdfOutputIntentInfo> outputIntents,
        PdfXmpMetadataInfo? xmpMetadata,
        PdfTaggedContentInfo? taggedContent,
        PdfOptionalContentProperties? optionalContent,
        PdfDocumentOpenAction? openAction,
        PdfViewerPreferences? viewerPreferences,
        IReadOnlyList<PdfFormField> formFields,
        string? acroFormDefaultAppearance,
        int? acroFormQuadding,
        PdfAcroFormXfaInfo? acroFormXfa,
        bool? acroFormNeedAppearances,
        int? acroFormSignatureFlags,
        PdfDocumentSecurityInfo security,
        string? catalogPageMode,
        string? catalogPageLayout,
        string? catalogVersion,
        string? catalogLanguage) {
        Metadata = metadata;
        Pages = pages;
        Outlines = outlines;
        PageLabels = pageLabels;
        NamedDestinations = namedDestinations;
        CatalogActions = catalogActions;
        Attachments = attachments;
        OutputIntents = outputIntents;
        XmpMetadata = xmpMetadata;
        TaggedContent = taggedContent;
        OptionalContent = optionalContent;
        OpenAction = openAction;
        ViewerPreferences = viewerPreferences;
        FormFields = formFields;
        AcroFormDefaultAppearance = acroFormDefaultAppearance;
        AcroFormQuadding = acroFormQuadding;
        AcroFormXfa = acroFormXfa;
        AcroFormNeedAppearances = acroFormNeedAppearances;
        AcroFormSignatureFlags = acroFormSignatureFlags;
        Security = security;
        CatalogPageMode = catalogPageMode;
        CatalogPageLayout = catalogPageLayout;
        CatalogVersion = catalogVersion;
        CatalogLanguage = catalogLanguage;
    }

    /// <summary>Document metadata read from the PDF Info dictionary when available.</summary>
    public PdfMetadata Metadata { get; }

    /// <summary>Logical pages in document order.</summary>
    public IReadOnlyList<PdfLogicalPage> Pages { get; }

    /// <summary>Logical pages grouped by one-based source page number. Range-based loads can contain the same source page more than once.</summary>
    public IReadOnlyDictionary<int, IReadOnlyList<PdfLogicalPage>> PagesBySourcePageNumber {
        get {
            if (_pagesBySourcePageNumber is not null) {
                return _pagesBySourcePageNumber;
            }

            var grouped = new Dictionary<int, List<PdfLogicalPage>>();
            for (int i = 0; i < Pages.Count; i++) {
                PdfLogicalPage page = Pages[i];
                if (!grouped.TryGetValue(page.PageNumber, out List<PdfLogicalPage>? pages)) {
                    pages = new List<PdfLogicalPage>();
                    grouped.Add(page.PageNumber, pages);
                }

                pages.Add(page);
            }

            _pagesBySourcePageNumber = ToReadOnlyLookup(grouped);
            return _pagesBySourcePageNumber;
        }
    }

    /// <summary>Top-level document outline/bookmark entries.</summary>
    public IReadOnlyList<PdfOutlineItem> Outlines { get; }

    /// <summary>Page-label rules discovered from the document catalog.</summary>
    public IReadOnlyList<PdfPageLabel> PageLabels { get; }

    /// <summary>Named destinations discovered from the document catalog.</summary>
    public IReadOnlyList<PdfNamedDestination> NamedDestinations { get; }

    /// <summary>Simple document open action discovered from the document catalog, when supported.</summary>
    public PdfDocumentOpenAction? OpenAction { get; }

    /// <summary>Simple viewer preference entries discovered from the document catalog, when supported.</summary>
    public PdfViewerPreferences? ViewerPreferences { get; }

    /// <summary>Simple AcroForm fields discovered from the document catalog.</summary>
    public IReadOnlyList<PdfFormField> FormFields { get; }

    /// <summary>AcroForm default appearance string from /DA, when present.</summary>
    public string? AcroFormDefaultAppearance { get; }

    /// <summary>True when an AcroForm default appearance string was readable.</summary>
    public bool HasAcroFormDefaultAppearance => !string.IsNullOrEmpty(AcroFormDefaultAppearance);

    /// <summary>Raw AcroForm default /Q quadding value, when present.</summary>
    public int? AcroFormQuadding { get; }

    /// <summary>True when an AcroForm default /Q quadding value was readable.</summary>
    public bool HasAcroFormQuadding => AcroFormQuadding.HasValue;

    /// <summary>Common AcroForm default text alignment inferred from /Q quadding.</summary>
    public PdfFormFieldTextAlignment AcroFormTextAlignment => ToTextAlignment(AcroFormQuadding);

    /// <summary>AcroForm XFA packet metadata when /AcroForm /XFA is present.</summary>
    public PdfAcroFormXfaInfo? AcroFormXfa { get; }

    /// <summary>True when the AcroForm contains an XFA packet entry.</summary>
    public bool HasAcroFormXfa => AcroFormXfa is not null;

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

    /// <summary>True when AcroForm /SigFlags indicates that the document contains signatures.</summary>
    public bool AcroFormSignaturesExist => HasAcroFormSignatureFlag(AcroFormSignaturesExistFlag);

    /// <summary>True when AcroForm /SigFlags indicates that the document should only be saved by appending changes.</summary>
    public bool AcroFormAppendOnly => HasAcroFormSignatureFlag(AcroFormAppendOnlyFlag);

    /// <summary>Named simple AcroForm fields keyed by fully qualified field name.</summary>
    public IReadOnlyDictionary<string, PdfFormField> FormFieldsByName {
        get {
            if (_formFieldsByName is not null) {
                return _formFieldsByName;
            }

            var fields = new Dictionary<string, PdfFormField>(StringComparer.Ordinal);
            for (int i = 0; i < FormFields.Count; i++) {
                PdfFormField formField = FormFields[i];
                string? name = formField.Name;
                if (name is not null && name.Length > 0 && !fields.ContainsKey(name)) {
                    fields.Add(name, formField);
                }
            }

            _formFieldsByName = new System.Collections.ObjectModel.ReadOnlyDictionary<string, PdfFormField>(fields);
            return _formFieldsByName;
        }
    }

    /// <summary>Fully qualified names for simple AcroForm fields that have a readable name.</summary>
    public IReadOnlyList<string> FormFieldNames {
        get {
            if (_formFieldNames is not null) {
                return _formFieldNames;
            }

            _formFieldNames = FormFieldsByName.Keys.ToArray();
            return _formFieldNames;
        }
    }

    /// <summary>Simple AcroForm fields grouped by common field kind.</summary>
    public IReadOnlyDictionary<PdfFormFieldKind, IReadOnlyList<PdfFormField>> FormFieldsByKind {
        get {
            if (_formFieldsByKind is not null) {
                return _formFieldsByKind;
            }

            var grouped = new Dictionary<PdfFormFieldKind, List<PdfFormField>>();
            for (int i = 0; i < FormFields.Count; i++) {
                PdfFormField formField = FormFields[i];
                if (!grouped.TryGetValue(formField.Kind, out List<PdfFormField>? fields)) {
                    fields = new List<PdfFormField>();
                    grouped.Add(formField.Kind, fields);
                }

                fields.Add(formField);
            }

            var result = new Dictionary<PdfFormFieldKind, IReadOnlyList<PdfFormField>>();
            foreach (var item in grouped) {
                result.Add(item.Key, item.Value.AsReadOnly());
            }

            _formFieldsByKind = new System.Collections.ObjectModel.ReadOnlyDictionary<PdfFormFieldKind, IReadOnlyList<PdfFormField>>(result);
            return _formFieldsByKind;
        }
    }

    /// <summary>Simple AcroForm fields grouped by one-based page number for fields that have readable widgets.</summary>
    public IReadOnlyDictionary<int, IReadOnlyList<PdfFormField>> FormFieldsByPageNumber {
        get {
            if (_formFieldsByPageNumber is not null) {
                return _formFieldsByPageNumber;
            }

            var grouped = new Dictionary<int, List<PdfFormField>>();
            IReadOnlyList<PdfLogicalFormWidget> widgets = FormWidgets;
            for (int i = 0; i < widgets.Count; i++) {
                PdfLogicalFormWidget widget = widgets[i];
                if (!grouped.TryGetValue(widget.PageNumber, out List<PdfFormField>? pageFields)) {
                    pageFields = new List<PdfFormField>();
                    grouped.Add(widget.PageNumber, pageFields);
                }

                if (!pageFields.Contains(widget.Field)) {
                    pageFields.Add(widget.Field);
                }
            }

            var result = new Dictionary<int, IReadOnlyList<PdfFormField>>();
            foreach (var item in grouped) {
                result.Add(item.Key, item.Value.AsReadOnly());
            }

            _formFieldsByPageNumber = new System.Collections.ObjectModel.ReadOnlyDictionary<int, IReadOnlyList<PdfFormField>>(result);
            return _formFieldsByPageNumber;
        }
    }

    /// <summary>All line-level text blocks flattened in page order.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> TextBlocks {
        get {
            _textBlocks ??= FlattenPageItems(Pages, page => page.TextBlocks);
            return _textBlocks;
        }
    }

    /// <summary>All heuristic heading objects flattened in page order.</summary>
    public IReadOnlyList<PdfLogicalHeading> Headings {
        get {
            _headings ??= FlattenPageItems(Pages, page => page.Headings);
            return _headings;
        }
    }

    /// <summary>All heuristic paragraph objects flattened in page order.</summary>
    public IReadOnlyList<PdfLogicalParagraph> Paragraphs {
        get {
            _paragraphs ??= FlattenPageItems(Pages, page => page.Paragraphs);
            return _paragraphs;
        }
    }

    /// <summary>All detected list item objects flattened in page order.</summary>
    public IReadOnlyList<PdfLogicalListItem> ListItems {
        get {
            _listItems ??= FlattenPageItems(Pages, page => page.ListItems);
            return _listItems;
        }
    }

    /// <summary>All detected table objects flattened in page order.</summary>
    public IReadOnlyList<PdfLogicalTable> Tables {
        get {
            _tables ??= FlattenPageItems(Pages, page => page.Tables);
            return _tables;
        }
    }

    /// <summary>All image XObject entries flattened in page order.</summary>
    public IReadOnlyList<PdfLogicalImage> Images {
        get {
            _images ??= FlattenPageItems(Pages, page => page.Images);
            return _images;
        }
    }

    /// <summary>All URI, named-destination, direct-destination, named-action, and remote GoTo link annotations flattened in page order.</summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> Links {
        get {
            if (_links is not null) {
                return _links;
            }

            var links = new List<PdfLogicalLinkAnnotation>();
            for (int i = 0; i < Pages.Count; i++) {
                links.AddRange(Pages[i].Links);
            }

            _links = links.AsReadOnly();
            return _links;
        }
    }

    /// <summary>URI link annotations grouped by URI action target.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfLogicalLinkAnnotation>> LinksByUri {
        get {
            if (_linksByUri is not null) {
                return _linksByUri;
            }

            var grouped = new Dictionary<string, List<PdfLogicalLinkAnnotation>>(StringComparer.Ordinal);
            IReadOnlyList<PdfLogicalLinkAnnotation> links = Links;
            for (int i = 0; i < links.Count; i++) {
                PdfLogicalLinkAnnotation link = links[i];
                string? uri = link.Uri;
                if (uri is null || uri.Length == 0) {
                    continue;
                }

                if (!grouped.TryGetValue(uri, out List<PdfLogicalLinkAnnotation>? uriLinks)) {
                    uriLinks = new List<PdfLogicalLinkAnnotation>();
                    grouped.Add(uri, uriLinks);
                }

                uriLinks.Add(link);
            }

            _linksByUri = ToReadOnlyLookup(grouped);
            return _linksByUri;
        }
    }

    /// <summary>Internal named-destination link annotations grouped by destination name.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfLogicalLinkAnnotation>> LinksByDestinationName {
        get {
            if (_linksByDestinationName is not null) {
                return _linksByDestinationName;
            }

            var grouped = new Dictionary<string, List<PdfLogicalLinkAnnotation>>(StringComparer.Ordinal);
            IReadOnlyList<PdfLogicalLinkAnnotation> links = Links;
            for (int i = 0; i < links.Count; i++) {
                PdfLogicalLinkAnnotation link = links[i];
                string? destinationName = link.DestinationName;
                if (destinationName is null || destinationName.Length == 0) {
                    continue;
                }

                if (!grouped.TryGetValue(destinationName, out List<PdfLogicalLinkAnnotation>? destinationLinks)) {
                    destinationLinks = new List<PdfLogicalLinkAnnotation>();
                    grouped.Add(destinationName, destinationLinks);
                }

                destinationLinks.Add(link);
            }

            _linksByDestinationName = ToReadOnlyLookup(grouped);
            return _linksByDestinationName;
        }
    }

    /// <summary>Internal direct-destination link annotations grouped by one-based destination page number.</summary>
    public IReadOnlyDictionary<int, IReadOnlyList<PdfLogicalLinkAnnotation>> LinksByDestinationPageNumber {
        get {
            if (_linksByDestinationPageNumber is not null) {
                return _linksByDestinationPageNumber;
            }

            var grouped = new Dictionary<int, List<PdfLogicalLinkAnnotation>>();
            IReadOnlyList<PdfLogicalLinkAnnotation> links = Links;
            for (int i = 0; i < links.Count; i++) {
                PdfLogicalLinkAnnotation link = links[i];
                if (!link.DestinationPageNumber.HasValue) {
                    continue;
                }

                int destinationPageNumber = link.DestinationPageNumber.Value;
                if (!grouped.TryGetValue(destinationPageNumber, out List<PdfLogicalLinkAnnotation>? destinationLinks)) {
                    destinationLinks = new List<PdfLogicalLinkAnnotation>();
                    grouped.Add(destinationPageNumber, destinationLinks);
                }

                destinationLinks.Add(link);
            }

            _linksByDestinationPageNumber = ToReadOnlyLookup(grouped);
            return _linksByDestinationPageNumber;
        }
    }

    /// <summary>Named-action link annotations grouped by viewer action name.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfLogicalLinkAnnotation>> LinksByNamedAction {
        get {
            if (_linksByNamedAction is not null) {
                return _linksByNamedAction;
            }

            var grouped = new Dictionary<string, List<PdfLogicalLinkAnnotation>>(StringComparer.Ordinal);
            IReadOnlyList<PdfLogicalLinkAnnotation> links = Links;
            for (int i = 0; i < links.Count; i++) {
                PdfLogicalLinkAnnotation link = links[i];
                string? namedAction = link.NamedAction;
                if (namedAction is null || namedAction.Length == 0) {
                    continue;
                }

                if (!grouped.TryGetValue(namedAction, out List<PdfLogicalLinkAnnotation>? actionLinks)) {
                    actionLinks = new List<PdfLogicalLinkAnnotation>();
                    grouped.Add(namedAction, actionLinks);
                }

                actionLinks.Add(link);
            }

            _linksByNamedAction = ToReadOnlyLookup(grouped);
            return _linksByNamedAction;
        }
    }

    /// <summary>Remote GoTo link annotations grouped by target file.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfLogicalLinkAnnotation>> LinksByRemoteFile {
        get {
            if (_linksByRemoteFile is not null) {
                return _linksByRemoteFile;
            }

            var grouped = new Dictionary<string, List<PdfLogicalLinkAnnotation>>(StringComparer.Ordinal);
            IReadOnlyList<PdfLogicalLinkAnnotation> links = Links;
            for (int i = 0; i < links.Count; i++) {
                PdfLogicalLinkAnnotation link = links[i];
                string? remoteFile = link.RemoteFile;
                if (remoteFile is null || remoteFile.Length == 0) {
                    continue;
                }

                if (!grouped.TryGetValue(remoteFile, out List<PdfLogicalLinkAnnotation>? fileLinks)) {
                    fileLinks = new List<PdfLogicalLinkAnnotation>();
                    grouped.Add(remoteFile, fileLinks);
                }

                fileLinks.Add(link);
            }

            _linksByRemoteFile = ToReadOnlyLookup(grouped);
            return _linksByRemoteFile;
        }
    }

    /// <summary>All AcroForm widget annotations flattened in page order.</summary>
    public IReadOnlyList<PdfLogicalFormWidget> FormWidgets {
        get {
            if (_formWidgets is not null) {
                return _formWidgets;
            }

            var widgets = new List<PdfLogicalFormWidget>();
            for (int i = 0; i < Pages.Count; i++) {
                widgets.AddRange(Pages[i].FormWidgets);
            }

            _formWidgets = widgets.AsReadOnly();
            return _formWidgets;
        }
    }

    /// <summary>AcroForm widget annotations grouped by fully qualified field name.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfLogicalFormWidget>> FormWidgetsByFieldName {
        get {
            if (_formWidgetsByFieldName is not null) {
                return _formWidgetsByFieldName;
            }

            var grouped = new Dictionary<string, List<PdfLogicalFormWidget>>(StringComparer.Ordinal);
            IReadOnlyList<PdfLogicalFormWidget> widgets = FormWidgets;
            for (int i = 0; i < widgets.Count; i++) {
                PdfLogicalFormWidget widget = widgets[i];
                string? fieldName = widget.FieldName;
                if (fieldName is null || fieldName.Length == 0) {
                    continue;
                }

                if (!grouped.TryGetValue(fieldName, out List<PdfLogicalFormWidget>? fieldWidgets)) {
                    fieldWidgets = new List<PdfLogicalFormWidget>();
                    grouped.Add(fieldName, fieldWidgets);
                }

                fieldWidgets.Add(widget);
            }

            var result = new Dictionary<string, IReadOnlyList<PdfLogicalFormWidget>>(StringComparer.Ordinal);
            foreach (var item in grouped) {
                result.Add(item.Key, item.Value.AsReadOnly());
            }

            _formWidgetsByFieldName = new System.Collections.ObjectModel.ReadOnlyDictionary<string, IReadOnlyList<PdfLogicalFormWidget>>(result);
            return _formWidgetsByFieldName;
        }
    }

    /// <summary>AcroForm widget annotations grouped by one-based page number.</summary>
    public IReadOnlyDictionary<int, IReadOnlyList<PdfLogicalFormWidget>> FormWidgetsByPageNumber {
        get {
            if (_formWidgetsByPageNumber is not null) {
                return _formWidgetsByPageNumber;
            }

            var grouped = new Dictionary<int, List<PdfLogicalFormWidget>>();
            for (int i = 0; i < Pages.Count; i++) {
                PdfLogicalPage page = Pages[i];
                if (page.FormWidgets.Count == 0) {
                    continue;
                }

                if (!grouped.TryGetValue(page.PageNumber, out List<PdfLogicalFormWidget>? pageWidgets)) {
                    pageWidgets = new List<PdfLogicalFormWidget>();
                    grouped.Add(page.PageNumber, pageWidgets);
                }

                pageWidgets.AddRange(page.FormWidgets);
            }

            var result = new Dictionary<int, IReadOnlyList<PdfLogicalFormWidget>>();
            foreach (var item in grouped) {
                result.Add(item.Key, item.Value.AsReadOnly());
            }

            _formWidgetsByPageNumber = new System.Collections.ObjectModel.ReadOnlyDictionary<int, IReadOnlyList<PdfLogicalFormWidget>>(result);
            return _formWidgetsByPageNumber;
        }
    }

    /// <summary>Catalog page mode, for example UseOutlines or FullScreen, when present.</summary>
    public string? CatalogPageMode { get; }

    /// <summary>Catalog page layout, for example SinglePage or TwoColumnLeft, when present.</summary>
    public string? CatalogPageLayout { get; }

    /// <summary>Catalog PDF version override, for example 1.7, when present.</summary>
    public string? CatalogVersion { get; }

    /// <summary>Catalog language tag, for example en-US or pl-PL, when present.</summary>
    public string? CatalogLanguage { get; }

    /// <summary>Number of pages in the logical document.</summary>
    public int PageCount => Pages.Count;

    /// <summary>True when at least one logical page for the one-based source page number is present.</summary>
    public bool HasSourcePage(int pageNumber) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        return PagesBySourcePageNumber.ContainsKey(pageNumber);
    }

    /// <summary>True when at least one outline/bookmark entry was read from the catalog.</summary>
    public bool HasOutlines => Outlines.Count > 0;

    /// <summary>True when at least one readable page-label rule was read from the catalog.</summary>
    public bool HasReadablePageLabels => PageLabels.Count > 0;

    /// <summary>True when at least one named destination was read from the catalog.</summary>
    public bool HasNamedDestinations => NamedDestinations.Count > 0;

    /// <summary>True when a simple document open action was read from the catalog.</summary>
    public bool HasReadableOpenAction => OpenAction is not null;

    /// <summary>True when simple viewer preferences were read from the catalog.</summary>
    public bool HasReadableViewerPreferences => ViewerPreferences is not null;

    /// <summary>True when at least one URI, named-destination, direct-destination, named-action, or remote GoTo link annotation was placed on a logical page.</summary>
    public bool HasLinks => Links.Count > 0;

    /// <summary>True when at least one simple AcroForm field was read from the document catalog.</summary>
    public bool HasFormFields => FormFields.Count > 0;

    /// <summary>True when at least one AcroForm widget annotation was placed on a logical page.</summary>
    public bool HasFormWidgets => FormWidgets.Count > 0;

    /// <summary>True when at least one logical element of the requested kind is present.</summary>
    public bool HasElementKind(PdfLogicalElementKind kind) {
        return ElementsByKind.ContainsKey(kind);
    }

    /// <summary>Returns logical pages for a one-based source page number, preserving range-selection duplicates.</summary>
    public IReadOnlyList<PdfLogicalPage> GetPages(int pageNumber) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        return PagesBySourcePageNumber.TryGetValue(pageNumber, out IReadOnlyList<PdfLogicalPage>? pages)
            ? pages
            : Array.Empty<PdfLogicalPage>();
    }

    /// <summary>Attempts to get a simple AcroForm field by its fully qualified field name.</summary>
    public bool TryGetFormField(string name, out PdfFormField? field) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        return FormFieldsByName.TryGetValue(name, out field);
    }

    /// <summary>Returns simple AcroForm fields for the requested common field kind.</summary>
    public IReadOnlyList<PdfFormField> GetFormFields(PdfFormFieldKind kind) {
        return FormFieldsByKind.TryGetValue(kind, out IReadOnlyList<PdfFormField>? fields)
            ? fields
            : Array.Empty<PdfFormField>();
    }

    /// <summary>Returns simple AcroForm fields represented by widgets on a one-based page number.</summary>
    public IReadOnlyList<PdfFormField> GetFormFields(int pageNumber) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        return FormFieldsByPageNumber.TryGetValue(pageNumber, out IReadOnlyList<PdfFormField>? fields)
            ? fields
            : Array.Empty<PdfFormField>();
    }

    /// <summary>Returns logical URI link annotations for a URI action target.</summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> GetLinksByUri(string uri) {
        Guard.UriAction(uri, nameof(uri));
        return LinksByUri.TryGetValue(uri, out IReadOnlyList<PdfLogicalLinkAnnotation>? links)
            ? links
            : Array.Empty<PdfLogicalLinkAnnotation>();
    }

    /// <summary>Returns logical internal link annotations for a named destination.</summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> GetLinksByDestinationName(string destinationName) {
        Guard.NotNullOrWhiteSpace(destinationName, nameof(destinationName));
        return LinksByDestinationName.TryGetValue(destinationName, out IReadOnlyList<PdfLogicalLinkAnnotation>? links)
            ? links
            : Array.Empty<PdfLogicalLinkAnnotation>();
    }

    /// <summary>Returns logical internal direct-destination link annotations for a one-based destination page number.</summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> GetLinksByDestinationPageNumber(int pageNumber) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        return LinksByDestinationPageNumber.TryGetValue(pageNumber, out IReadOnlyList<PdfLogicalLinkAnnotation>? links)
            ? links
            : Array.Empty<PdfLogicalLinkAnnotation>();
    }

    /// <summary>Returns logical named-action link annotations for a viewer action name.</summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> GetLinksByNamedAction(string namedAction) {
        Guard.NotNullOrWhiteSpace(namedAction, nameof(namedAction));
        return LinksByNamedAction.TryGetValue(namedAction, out IReadOnlyList<PdfLogicalLinkAnnotation>? links)
            ? links
            : Array.Empty<PdfLogicalLinkAnnotation>();
    }

    /// <summary>Returns logical remote GoTo link annotations for a target file.</summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> GetLinksByRemoteFile(string remoteFile) {
        Guard.NotNullOrWhiteSpace(remoteFile, nameof(remoteFile));
        return LinksByRemoteFile.TryGetValue(remoteFile, out IReadOnlyList<PdfLogicalLinkAnnotation>? links)
            ? links
            : Array.Empty<PdfLogicalLinkAnnotation>();
    }

    /// <summary>Returns logical widget annotations for a fully qualified form field name.</summary>
    public IReadOnlyList<PdfLogicalFormWidget> GetFormWidgets(string fieldName) {
        Guard.NotNullOrWhiteSpace(fieldName, nameof(fieldName));
        return FormWidgetsByFieldName.TryGetValue(fieldName, out IReadOnlyList<PdfLogicalFormWidget>? widgets)
            ? widgets
            : Array.Empty<PdfLogicalFormWidget>();
    }

    /// <summary>Returns logical widget annotations for a one-based page number.</summary>
    public IReadOnlyList<PdfLogicalFormWidget> GetFormWidgets(int pageNumber) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        return FormWidgetsByPageNumber.TryGetValue(pageNumber, out IReadOnlyList<PdfLogicalFormWidget>? widgets)
            ? widgets
            : Array.Empty<PdfLogicalFormWidget>();
    }

    /// <summary>Returns logical elements of the requested kind in document order.</summary>
    public IReadOnlyList<IPdfLogicalElement> GetElements(PdfLogicalElementKind kind) {
        return ElementsByKind.TryGetValue(kind, out IReadOnlyList<IPdfLogicalElement>? elements)
            ? elements
            : Array.Empty<IPdfLogicalElement>();
    }

    /// <summary>Returns logical elements for a one-based source page number.</summary>
    public IReadOnlyList<IPdfLogicalElement> GetElements(int pageNumber) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        return ElementsByPageNumber.TryGetValue(pageNumber, out IReadOnlyList<IPdfLogicalElement>? elements)
            ? elements
            : Array.Empty<IPdfLogicalElement>();
    }

    /// <summary>All logical page elements flattened in page order.</summary>
    public IReadOnlyList<IPdfLogicalElement> Elements {
        get {
            if (_elements is not null) {
                return _elements;
            }

            var elements = new List<IPdfLogicalElement>();
            for (int i = 0; i < Pages.Count; i++) {
                elements.AddRange(Pages[i].Elements);
            }

            _elements = elements.AsReadOnly();
            return _elements;
        }
    }

    /// <summary>Logical page elements grouped by element kind.</summary>
    public IReadOnlyDictionary<PdfLogicalElementKind, IReadOnlyList<IPdfLogicalElement>> ElementsByKind {
        get {
            if (_elementsByKind is not null) {
                return _elementsByKind;
            }

            var grouped = new Dictionary<PdfLogicalElementKind, List<IPdfLogicalElement>>();
            IReadOnlyList<IPdfLogicalElement> elements = Elements;
            for (int i = 0; i < elements.Count; i++) {
                IPdfLogicalElement element = elements[i];
                if (!grouped.TryGetValue(element.Kind, out List<IPdfLogicalElement>? kindElements)) {
                    kindElements = new List<IPdfLogicalElement>();
                    grouped.Add(element.Kind, kindElements);
                }

                kindElements.Add(element);
            }

            _elementsByKind = ToReadOnlyLookup(grouped);
            return _elementsByKind;
        }
    }

    /// <summary>Logical page elements grouped by one-based source page number.</summary>
    public IReadOnlyDictionary<int, IReadOnlyList<IPdfLogicalElement>> ElementsByPageNumber {
        get {
            if (_elementsByPageNumber is not null) {
                return _elementsByPageNumber;
            }

            var grouped = new Dictionary<int, List<IPdfLogicalElement>>();
            for (int i = 0; i < Pages.Count; i++) {
                PdfLogicalPage page = Pages[i];
                if (!grouped.TryGetValue(page.PageNumber, out List<IPdfLogicalElement>? pageElements)) {
                    pageElements = new List<IPdfLogicalElement>();
                    grouped.Add(page.PageNumber, pageElements);
                }

                pageElements.AddRange(page.Elements);
            }

            _elementsByPageNumber = ToReadOnlyLookup(grouped);
            return _elementsByPageNumber;
        }
    }

    private static System.Collections.ObjectModel.ReadOnlyDictionary<string, IReadOnlyList<T>> ToReadOnlyLookup<T>(Dictionary<string, List<T>> grouped) {
        var result = new Dictionary<string, IReadOnlyList<T>>(StringComparer.Ordinal);
        foreach (var item in grouped) {
            result.Add(item.Key, item.Value.AsReadOnly());
        }

        return new System.Collections.ObjectModel.ReadOnlyDictionary<string, IReadOnlyList<T>>(result);
    }

    private static System.Collections.ObjectModel.ReadOnlyDictionary<TKey, IReadOnlyList<T>> ToReadOnlyLookup<TKey, T>(Dictionary<TKey, List<T>> grouped) where TKey : notnull {
        var result = new Dictionary<TKey, IReadOnlyList<T>>();
        foreach (var item in grouped) {
            result.Add(item.Key, item.Value.AsReadOnly());
        }

        return new System.Collections.ObjectModel.ReadOnlyDictionary<TKey, IReadOnlyList<T>>(result);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<T> FlattenPageItems<T>(IReadOnlyList<PdfLogicalPage> pages, Func<PdfLogicalPage, IReadOnlyList<T>> selector) {
        var items = new List<T>();
        for (int i = 0; i < pages.Count; i++) {
            items.AddRange(selector(pages[i]));
        }

        return items.AsReadOnly();
    }

    private bool HasAcroFormSignatureFlag(int flag) {
        return AcroFormSignatureFlags.HasValue && (AcroFormSignatureFlags.Value & flag) != 0;
    }

    /// <summary>Loads a PDF from bytes and returns the logical read model.</summary>
    public static PdfLogicalDocument Load(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return From(PdfReadDocument.Open(pdf), options);
    }

    /// <summary>Loads a PDF from bytes with explicit read limits or credentials and returns the logical read model.</summary>
    public static PdfLogicalDocument Load(byte[] pdf, PdfTextLayoutOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        return From(PdfReadDocument.Open(pdf, readOptions), options);
    }

    /// <summary>Loads a PDF from a file path and returns the logical read model.</summary>
    public static PdfLogicalDocument Load(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return From(PdfReadDocument.Open(path), options);
    }

    /// <summary>Loads a PDF from the current position of a readable stream and returns the logical read model.</summary>
    public static PdfLogicalDocument Load(Stream stream, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        return From(PdfReadDocument.Open(stream), options);
    }

    /// <summary>Loads selected source page ranges from PDF bytes into the logical read model, preserving caller order and overlaps.</summary>
    public static PdfLogicalDocument LoadPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return LoadPageRanges(pdf, null, pageRanges);
    }

    /// <summary>Loads selected source page ranges from PDF bytes into the logical read model, preserving caller order and overlaps.</summary>
    public static PdfLogicalDocument LoadPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return FromPageRanges(PdfReadDocument.Open(pdf), options, pageRanges);
    }

    /// <summary>Loads selected source page ranges from a file path into the logical read model, preserving caller order and overlaps.</summary>
    public static PdfLogicalDocument LoadPageRanges(string path, params PdfPageRange[] pageRanges) {
        return LoadPageRanges(path, null, pageRanges);
    }

    /// <summary>Loads selected source page ranges from a file path into the logical read model, preserving caller order and overlaps.</summary>
    public static PdfLogicalDocument LoadPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return FromPageRanges(PdfReadDocument.Open(path), options, pageRanges);
    }

    /// <summary>Loads selected source page ranges from the current position of a readable stream into the logical read model, preserving caller order and overlaps.</summary>
    public static PdfLogicalDocument LoadPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return LoadPageRanges(stream, null, pageRanges);
    }

    /// <summary>Loads selected source page ranges from the current position of a readable stream into the logical read model, preserving caller order and overlaps.</summary>
    public static PdfLogicalDocument LoadPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(stream, nameof(stream));
        return FromPageRanges(PdfReadDocument.Open(stream), options, pageRanges);
    }

    /// <summary>Builds the logical read model from an already parsed PDF document.</summary>
    public static PdfLogicalDocument From(PdfReadDocument document, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(document, nameof(document));

        var pageNumbers = new int[document.Pages.Count];
        for (int i = 0; i < document.Pages.Count; i++) {
            pageNumbers[i] = i + 1;
        }

        return FromPageNumbers(document, options, pageNumbers);
    }

    /// <summary>Builds a logical read model for selected source page ranges from an already parsed PDF document, preserving caller order and overlaps.</summary>
    public static PdfLogicalDocument FromPageRanges(PdfReadDocument document, params PdfPageRange[] pageRanges) {
        return FromPageRanges(document, null, pageRanges);
    }

    /// <summary>Builds a logical read model for selected source page ranges from an already parsed PDF document, preserving caller order and overlaps.</summary>
    public static PdfLogicalDocument FromPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(document, nameof(document));
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));

        return FromPageNumbers(document, options, pageNumbers);
    }

    private static PdfLogicalDocument FromPageNumbers(PdfReadDocument document, PdfTextLayoutOptions? options, int[] pageNumbers) {
        bool useDocumentWideObjects = PdfPageRangeObjectFilter.ShouldUseDocumentWideObjects(document.Pages.Count, pageNumbers);
        IReadOnlyList<PdfFormField> formFields = useDocumentWideObjects
            ? document.FormFields
            : PdfPageRangeObjectFilter.FilterFormFieldsByPageNumbers(document.FormFields, pageNumbers, preservePageDuplicates: false);
        IReadOnlyList<PdfOutlineItem> outlines = useDocumentWideObjects
            ? document.Outlines
            : PdfPageRangeObjectFilter.FilterOutlinesByPageNumbers(document.Outlines, pageNumbers);
        IReadOnlyList<PdfPageLabel> pageLabels = useDocumentWideObjects
            ? document.PageLabels
            : PdfPageRangeObjectFilter.FilterPageLabelsByPageNumbers(document.PageLabels, pageNumbers);
        IReadOnlyList<PdfNamedDestination> namedDestinations = useDocumentWideObjects
            ? document.NamedDestinations
            : PdfPageRangeObjectFilter.FilterNamedDestinationsByPageNumbers(document.NamedDestinations, pageNumbers);
        IReadOnlyList<PdfCatalogAction> catalogActions = useDocumentWideObjects
            ? document.CatalogActions
            : Array.Empty<PdfCatalogAction>();
        IReadOnlyList<PdfAttachmentInfo> attachments = useDocumentWideObjects
            ? document.Attachments
            : Array.Empty<PdfAttachmentInfo>();
        IReadOnlyList<PdfOutputIntentInfo> outputIntents = useDocumentWideObjects
            ? document.OutputIntents
            : Array.Empty<PdfOutputIntentInfo>();
        PdfXmpMetadataInfo? xmpMetadata = useDocumentWideObjects
            ? document.XmpMetadata
            : null;
        PdfTaggedContentInfo? taggedContent = useDocumentWideObjects
            ? document.TaggedContent
            : null;
        PdfOptionalContentProperties? optionalContent = useDocumentWideObjects
            ? document.OptionalContent
            : null;
        PdfDocumentOpenAction? openAction = useDocumentWideObjects
            ? document.OpenAction
            : PdfPageRangeObjectFilter.FilterOpenActionByPageNumbers(document.OpenAction, pageNumbers);

        var pages = new List<PdfLogicalPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            pages.Add(PdfLogicalPage.From(document, document.Pages[pageNumber - 1], pageNumber, options, formFields));
        }

        return new PdfLogicalDocument(
            document.Metadata,
            pages.AsReadOnly(),
            outlines,
            pageLabels,
            namedDestinations,
            catalogActions,
            attachments,
            outputIntents,
            xmpMetadata,
            taggedContent,
            optionalContent,
            openAction,
            document.ViewerPreferences,
            formFields,
            document.AcroFormDefaultAppearance,
            document.AcroFormQuadding,
            document.AcroFormXfa,
            document.AcroFormNeedAppearances,
            document.AcroFormSignatureFlags,
            document.Security,
            document.CatalogPageMode,
            document.CatalogPageLayout,
            document.CatalogVersion,
            document.CatalogLanguage);
    }

    private static PdfFormFieldTextAlignment ToTextAlignment(int? quadding) {
        switch (quadding) {
            case 0:
                return PdfFormFieldTextAlignment.Left;
            case 1:
                return PdfFormFieldTextAlignment.Center;
            case 2:
                return PdfFormFieldTextAlignment.Right;
            default:
                return PdfFormFieldTextAlignment.Unknown;
        }
    }
}

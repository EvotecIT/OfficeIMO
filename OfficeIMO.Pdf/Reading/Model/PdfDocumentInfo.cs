namespace OfficeIMO.Pdf;

/// <summary>
/// Basic document-level information useful for inspection and automation scenarios.
/// </summary>
public sealed partial class PdfDocumentInfo {
    private const int AcroFormSignaturesExistFlag = 1;
    private const int AcroFormAppendOnlyFlag = 2;
    private IReadOnlyList<PdfAnnotation>? _annotations;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfAnnotation>>? _annotationsBySubtype;
    private IReadOnlyList<string>? _annotationActionTypes;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfAnnotation>>? _annotationsByActionType;
    private IReadOnlyList<PdfLinkAnnotation>? _linkAnnotations;
    private IReadOnlyList<string>? _linkUris;
    private IReadOnlyList<string>? _linkDestinationNames;
    private IReadOnlyList<int>? _linkDestinationPageNumbers;
    private IReadOnlyList<string>? _linkNamedActions;
    private IReadOnlyList<string>? _linkRemoteFiles;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfLinkAnnotation>>? _linkAnnotationsByUri;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfLinkAnnotation>>? _linkAnnotationsByDestinationName;
    private IReadOnlyDictionary<int, IReadOnlyList<PdfLinkAnnotation>>? _linkAnnotationsByDestinationPageNumber;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfLinkAnnotation>>? _linkAnnotationsByNamedAction;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfLinkAnnotation>>? _linkAnnotationsByRemoteFile;
    private IReadOnlyList<string>? _namedDestinationNames;
    private IReadOnlyList<string>? _formFieldNames;
    private IReadOnlyDictionary<string, PdfFormField>? _formFieldsByName;
    private IReadOnlyDictionary<PdfFormFieldKind, IReadOnlyList<PdfFormField>>? _formFieldsByKind;
    private IReadOnlyDictionary<int, IReadOnlyList<PdfFormField>>? _formFieldsByPageNumber;
    private IReadOnlyList<PdfFormWidget>? _formWidgets;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfFormWidget>>? _formWidgetsByFieldName;
    private IReadOnlyDictionary<int, IReadOnlyList<PdfFormWidget>>? _formWidgetsByPageNumber;

    internal PdfDocumentInfo(IReadOnlyList<PdfPageInfo> pages, PdfMetadata metadata, IReadOnlyList<PdfOutlineItem> outlines, IReadOnlyList<PdfPageLabel> pageLabels, IReadOnlyList<PdfNamedDestination> namedDestinations, IReadOnlyList<PdfCatalogAction> catalogActions, IReadOnlyList<PdfAttachmentInfo> attachments, IReadOnlyList<PdfOutputIntentInfo> outputIntents, PdfXmpMetadataInfo? xmpMetadata, PdfTaggedContentInfo? taggedContent, PdfOptionalContentProperties? optionalContent, PdfDocumentOpenAction? openAction, PdfViewerPreferences? viewerPreferences, IReadOnlyList<PdfFormField> formFields, string? acroFormDefaultAppearance, int? acroFormQuadding, PdfAcroFormXfaInfo? acroFormXfa, bool? acroFormNeedAppearances, int? acroFormSignatureFlags, PdfDocumentSecurityInfo security, string? headerVersion, string? catalogPageMode, string? catalogPageLayout, string? catalogVersion, string? catalogLanguage, bool hasSignatures, bool hasForms, bool hasAnnotations, bool hasOutlines, bool hasCatalogViewSettings, bool hasPageLabels, bool hasCatalogNameTrees, bool hasNamedDestinations, bool hasOpenActions, bool hasViewerPreferences, bool hasTaggedContent, bool hasXmpMetadata, bool hasCatalogUri, bool hasOutputIntents, bool hasEmbeddedFiles, bool hasOptionalContent, bool hasActiveContent) {
        Pages = pages;
        Metadata = metadata;
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

    /// <summary>Number of generic page annotations read from all pages.</summary>
    public int AnnotationCount => Annotations.Count;

    /// <summary>Number of distinct primary or additional annotation action types read from all pages.</summary>
    public int AnnotationActionTypeCount => AnnotationActionTypes.Count;

    /// <summary>Number of simple link annotations read from all pages.</summary>
    public int LinkAnnotationCount => LinkAnnotations.Count;

    /// <summary>Generic page annotations read from all pages in document order.</summary>
    public IReadOnlyList<PdfAnnotation> Annotations {
        get {
            if (_annotations is not null) {
                return _annotations;
            }

            var annotations = new List<PdfAnnotation>();
            for (int i = 0; i < Pages.Count; i++) {
                for (int j = 0; j < Pages[i].Annotations.Count; j++) {
                    var annotation = Pages[i].Annotations[j];
                    annotations.Add(annotation.PageNumber.HasValue ? annotation : annotation.WithPageNumber(Pages[i].PageNumber));
                }
            }

            _annotations = annotations.AsReadOnly();
            return _annotations;
        }
    }

    /// <summary>Generic page annotations grouped by PDF annotation subtype name.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfAnnotation>> AnnotationsBySubtype {
        get {
            if (_annotationsBySubtype is not null) {
                return _annotationsBySubtype;
            }

            var grouped = new Dictionary<string, List<PdfAnnotation>>(StringComparer.Ordinal);
            foreach (var annotation in Annotations) {
                if (!grouped.TryGetValue(annotation.Subtype, out List<PdfAnnotation>? annotations)) {
                    annotations = new List<PdfAnnotation>();
                    grouped.Add(annotation.Subtype, annotations);
                }

                annotations.Add(annotation);
            }

            var result = new Dictionary<string, IReadOnlyList<PdfAnnotation>>(StringComparer.Ordinal);
            foreach (var item in grouped) {
                result.Add(item.Key, item.Value.AsReadOnly());
            }

            _annotationsBySubtype = new System.Collections.ObjectModel.ReadOnlyDictionary<string, IReadOnlyList<PdfAnnotation>>(result);
            return _annotationsBySubtype;
        }
    }

    /// <summary>Distinct primary and additional annotation action types read from all pages in first-seen document order.</summary>
    public IReadOnlyList<string> AnnotationActionTypes {
        get {
            if (_annotationActionTypes is not null) {
                return _annotationActionTypes;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var actionTypes = new List<string>();
            foreach (var annotation in Annotations) {
                if (!string.IsNullOrEmpty(annotation.ActionType) && seen.Add(annotation.ActionType!)) {
                    actionTypes.Add(annotation.ActionType!);
                }

                for (int i = 0; i < annotation.AdditionalActions.Count; i++) {
                    string actionType = annotation.AdditionalActions[i].ActionType;
                    if (seen.Add(actionType)) {
                        actionTypes.Add(actionType);
                    }
                }

                for (int i = 0; i < annotation.ChainedActions.Count; i++) {
                    string actionType = annotation.ChainedActions[i].ActionType;
                    if (seen.Add(actionType)) {
                        actionTypes.Add(actionType);
                    }
                }
            }

            _annotationActionTypes = actionTypes.AsReadOnly();
            return _annotationActionTypes;
        }
    }

    /// <summary>Generic page annotations grouped by primary or additional action type.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfAnnotation>> AnnotationsByActionType {
        get {
            if (_annotationsByActionType is not null) {
                return _annotationsByActionType;
            }

            var grouped = new Dictionary<string, List<PdfAnnotation>>(StringComparer.Ordinal);
            foreach (var annotation in Annotations) {
                if (!string.IsNullOrEmpty(annotation.ActionType)) {
                    AddAnnotation(grouped, annotation.ActionType!, annotation);
                }

                for (int i = 0; i < annotation.AdditionalActions.Count; i++) {
                    AddAnnotation(grouped, annotation.AdditionalActions[i].ActionType, annotation);
                }

                for (int i = 0; i < annotation.ChainedActions.Count; i++) {
                    AddAnnotation(grouped, annotation.ChainedActions[i].ActionType, annotation);
                }
            }

            _annotationsByActionType = ToReadOnlyLookup(grouped);
            return _annotationsByActionType;
        }
    }

    /// <summary>Number of distinct simple URI link targets read from all pages.</summary>
    public int LinkUriCount => LinkUris.Count;

    /// <summary>Number of distinct simple named-destination link targets read from all pages.</summary>
    public int LinkDestinationCount => LinkDestinationNames.Count;

    /// <summary>Number of distinct simple direct-destination link target pages read from all pages.</summary>
    public int LinkDestinationPageNumberCount => LinkDestinationPageNumbers.Count;

    /// <summary>Number of distinct simple named viewer actions read from all pages.</summary>
    public int LinkNamedActionCount => LinkNamedActions.Count;

    /// <summary>Number of distinct remote GoTo target files read from all pages.</summary>
    public int LinkRemoteFileCount => LinkRemoteFiles.Count;

    /// <summary>Number of named destinations read from the document catalog.</summary>
    public int NamedDestinationCount => NamedDestinations.Count;

    /// <summary>Number of simple AcroForm fields read from the document catalog.</summary>
    public int FormFieldCount => FormFields.Count;

    /// <summary>Number of simple AcroForm widget annotations read from the document catalog fields.</summary>
    public int FormWidgetCount => FormWidgets.Count;

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

    /// <summary>Distinct simple direct-destination link target page numbers read from all pages in first-seen document order.</summary>
    public IReadOnlyList<int> LinkDestinationPageNumbers {
        get {
            if (_linkDestinationPageNumbers is not null) {
                return _linkDestinationPageNumbers;
            }

            var seen = new HashSet<int>();
            var pageNumbers = new List<int>();
            foreach (var link in LinkAnnotations) {
                if (link.DestinationPageNumber.HasValue && seen.Add(link.DestinationPageNumber.Value)) {
                    pageNumbers.Add(link.DestinationPageNumber.Value);
                }
            }

            _linkDestinationPageNumbers = pageNumbers.AsReadOnly();
            return _linkDestinationPageNumbers;
        }
    }

    /// <summary>Distinct simple named viewer actions read from all pages in first-seen document order.</summary>
    public IReadOnlyList<string> LinkNamedActions {
        get {
            if (_linkNamedActions is not null) {
                return _linkNamedActions;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var actions = new List<string>();
            foreach (var link in LinkAnnotations) {
                if (link.NamedAction != null && seen.Add(link.NamedAction)) {
                    actions.Add(link.NamedAction);
                }
            }

            _linkNamedActions = actions.AsReadOnly();
            return _linkNamedActions;
        }
    }

    /// <summary>Distinct remote GoTo target files read from all pages in first-seen document order.</summary>
    public IReadOnlyList<string> LinkRemoteFiles {
        get {
            if (_linkRemoteFiles is not null) {
                return _linkRemoteFiles;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var files = new List<string>();
            foreach (var link in LinkAnnotations) {
                if (link.RemoteFile != null && seen.Add(link.RemoteFile)) {
                    files.Add(link.RemoteFile);
                }
            }

            _linkRemoteFiles = files.AsReadOnly();
            return _linkRemoteFiles;
        }
    }

    /// <summary>Simple URI link annotations grouped by URI action target.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfLinkAnnotation>> LinkAnnotationsByUri {
        get {
            if (_linkAnnotationsByUri is not null) {
                return _linkAnnotationsByUri;
            }

            var grouped = new Dictionary<string, List<PdfLinkAnnotation>>(StringComparer.Ordinal);
            foreach (var link in LinkAnnotations) {
                string? uri = link.Uri;
                if (string.IsNullOrEmpty(uri)) {
                    continue;
                }

                if (!grouped.TryGetValue(uri!, out List<PdfLinkAnnotation>? links)) {
                    links = new List<PdfLinkAnnotation>();
                    grouped.Add(uri!, links);
                }

                links.Add(link);
            }

            _linkAnnotationsByUri = ToReadOnlyLookup(grouped);
            return _linkAnnotationsByUri;
        }
    }

    /// <summary>Simple named-destination link annotations grouped by destination name.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfLinkAnnotation>> LinkAnnotationsByDestinationName {
        get {
            if (_linkAnnotationsByDestinationName is not null) {
                return _linkAnnotationsByDestinationName;
            }

            var grouped = new Dictionary<string, List<PdfLinkAnnotation>>(StringComparer.Ordinal);
            foreach (var link in LinkAnnotations) {
                string? destinationName = link.DestinationName;
                if (string.IsNullOrEmpty(destinationName)) {
                    continue;
                }

                if (!grouped.TryGetValue(destinationName!, out List<PdfLinkAnnotation>? links)) {
                    links = new List<PdfLinkAnnotation>();
                    grouped.Add(destinationName!, links);
                }

                links.Add(link);
            }

            _linkAnnotationsByDestinationName = ToReadOnlyLookup(grouped);
            return _linkAnnotationsByDestinationName;
        }
    }

    /// <summary>Simple direct-destination link annotations grouped by one-based destination page number.</summary>
    public IReadOnlyDictionary<int, IReadOnlyList<PdfLinkAnnotation>> LinkAnnotationsByDestinationPageNumber {
        get {
            if (_linkAnnotationsByDestinationPageNumber is not null) {
                return _linkAnnotationsByDestinationPageNumber;
            }

            var grouped = new Dictionary<int, List<PdfLinkAnnotation>>();
            foreach (var link in LinkAnnotations) {
                if (!link.DestinationPageNumber.HasValue) {
                    continue;
                }

                int destinationPageNumber = link.DestinationPageNumber.Value;
                if (!grouped.TryGetValue(destinationPageNumber, out List<PdfLinkAnnotation>? links)) {
                    links = new List<PdfLinkAnnotation>();
                    grouped.Add(destinationPageNumber, links);
                }

                links.Add(link);
            }

            _linkAnnotationsByDestinationPageNumber = ToReadOnlyLookup(grouped);
            return _linkAnnotationsByDestinationPageNumber;
        }
    }

    /// <summary>Simple named-action link annotations grouped by viewer action name.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfLinkAnnotation>> LinkAnnotationsByNamedAction {
        get {
            if (_linkAnnotationsByNamedAction is not null) {
                return _linkAnnotationsByNamedAction;
            }

            var grouped = new Dictionary<string, List<PdfLinkAnnotation>>(StringComparer.Ordinal);
            foreach (var link in LinkAnnotations) {
                string? namedAction = link.NamedAction;
                if (string.IsNullOrEmpty(namedAction)) {
                    continue;
                }

                if (!grouped.TryGetValue(namedAction!, out List<PdfLinkAnnotation>? links)) {
                    links = new List<PdfLinkAnnotation>();
                    grouped.Add(namedAction!, links);
                }

                links.Add(link);
            }

            _linkAnnotationsByNamedAction = ToReadOnlyLookup(grouped);
            return _linkAnnotationsByNamedAction;
        }
    }

    /// <summary>Simple remote GoTo link annotations grouped by target file.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfLinkAnnotation>> LinkAnnotationsByRemoteFile {
        get {
            if (_linkAnnotationsByRemoteFile is not null) {
                return _linkAnnotationsByRemoteFile;
            }

            var grouped = new Dictionary<string, List<PdfLinkAnnotation>>(StringComparer.Ordinal);
            foreach (var link in LinkAnnotations) {
                string? remoteFile = link.RemoteFile;
                if (string.IsNullOrEmpty(remoteFile)) {
                    continue;
                }

                if (!grouped.TryGetValue(remoteFile!, out List<PdfLinkAnnotation>? links)) {
                    links = new List<PdfLinkAnnotation>();
                    grouped.Add(remoteFile!, links);
                }

                links.Add(link);
            }

            _linkAnnotationsByRemoteFile = ToReadOnlyLookup(grouped);
            return _linkAnnotationsByRemoteFile;
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

            _formFieldNames = FormFieldsByName.Keys.ToArray();
            return _formFieldNames;
        }
    }

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
            var memberships = new Dictionary<int, HashSet<PdfFormField>>();
            for (int i = 0; i < FormFields.Count; i++) {
                PdfFormField formField = FormFields[i];
                for (int j = 0; j < formField.Widgets.Count; j++) {
                    int? pageNumber = formField.Widgets[j].PageNumber;
                    if (!pageNumber.HasValue) {
                        continue;
                    }

                    if (!grouped.TryGetValue(pageNumber.Value, out List<PdfFormField>? fields)) {
                        fields = new List<PdfFormField>();
                        grouped.Add(pageNumber.Value, fields);
                        memberships.Add(pageNumber.Value, new HashSet<PdfFormField>());
                    }

                    if (memberships[pageNumber.Value].Add(formField)) {
                        fields.Add(formField);
                    }
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

    /// <summary>Simple AcroForm widget annotations flattened in field and widget order.</summary>
    public IReadOnlyList<PdfFormWidget> FormWidgets {
        get {
            if (_formWidgets is not null) {
                return _formWidgets;
            }

            var widgets = new List<PdfFormWidget>();
            for (int i = 0; i < FormFields.Count; i++) {
                widgets.AddRange(FormFields[i].Widgets);
            }

            _formWidgets = widgets.AsReadOnly();
            return _formWidgets;
        }
    }

    /// <summary>Simple AcroForm widget annotations grouped by fully qualified field name.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfFormWidget>> FormWidgetsByFieldName {
        get {
            if (_formWidgetsByFieldName is not null) {
                return _formWidgetsByFieldName;
            }

            var grouped = new Dictionary<string, List<PdfFormWidget>>(StringComparer.Ordinal);
            for (int i = 0; i < FormFields.Count; i++) {
                PdfFormField formField = FormFields[i];
                string? name = formField.Name;
                if (name is null || name.Length == 0 || formField.Widgets.Count == 0) {
                    continue;
                }

                if (!grouped.TryGetValue(name, out List<PdfFormWidget>? widgets)) {
                    widgets = new List<PdfFormWidget>();
                    grouped.Add(name, widgets);
                }

                widgets.AddRange(formField.Widgets);
            }

            var result = new Dictionary<string, IReadOnlyList<PdfFormWidget>>(StringComparer.Ordinal);
            foreach (var item in grouped) {
                result.Add(item.Key, item.Value.AsReadOnly());
            }

            _formWidgetsByFieldName = new System.Collections.ObjectModel.ReadOnlyDictionary<string, IReadOnlyList<PdfFormWidget>>(result);
            return _formWidgetsByFieldName;
        }
    }

    /// <summary>Simple AcroForm widget annotations grouped by one-based page number.</summary>
    public IReadOnlyDictionary<int, IReadOnlyList<PdfFormWidget>> FormWidgetsByPageNumber {
        get {
            if (_formWidgetsByPageNumber is not null) {
                return _formWidgetsByPageNumber;
            }

            var grouped = new Dictionary<int, List<PdfFormWidget>>();
            for (int i = 0; i < Pages.Count; i++) {
                PdfPageInfo page = Pages[i];
                if (page.FormWidgets.Count == 0) {
                    continue;
                }

                if (!grouped.TryGetValue(page.PageNumber, out List<PdfFormWidget>? widgets)) {
                    widgets = new List<PdfFormWidget>();
                    grouped.Add(page.PageNumber, widgets);
                }

                widgets.AddRange(page.FormWidgets);
            }

            var result = new Dictionary<int, IReadOnlyList<PdfFormWidget>>();
            foreach (var item in grouped) {
                result.Add(item.Key, item.Value.AsReadOnly());
            }

            _formWidgetsByPageNumber = new System.Collections.ObjectModel.ReadOnlyDictionary<int, IReadOnlyList<PdfFormWidget>>(result);
            return _formWidgetsByPageNumber;
        }
    }

    /// <summary>True when at least one simple link annotation was read from the document pages.</summary>
    public bool HasLinkAnnotations => LinkAnnotationCount > 0;

    /// <summary>Returns generic page annotations with a matching PDF annotation subtype name.</summary>
    public IReadOnlyList<PdfAnnotation> GetAnnotationsBySubtype(string subtype) {
        Guard.NotNullOrWhiteSpace(subtype, nameof(subtype));
        return AnnotationsBySubtype.TryGetValue(subtype, out IReadOnlyList<PdfAnnotation>? annotations)
            ? annotations
            : Array.Empty<PdfAnnotation>();
    }

    /// <summary>Returns generic page annotations with a matching primary or additional action type.</summary>
    public IReadOnlyList<PdfAnnotation> GetAnnotationsByActionType(string actionType) {
        Guard.NotNullOrWhiteSpace(actionType, nameof(actionType));
        return AnnotationsByActionType.TryGetValue(actionType, out IReadOnlyList<PdfAnnotation>? annotations)
            ? annotations
            : Array.Empty<PdfAnnotation>();
    }

    /// <summary>Returns simple URI link annotations for a URI action target.</summary>
    public IReadOnlyList<PdfLinkAnnotation> GetLinkAnnotationsByUri(string uri) {
        Guard.UriAction(uri, nameof(uri));
        return LinkAnnotationsByUri.TryGetValue(uri, out IReadOnlyList<PdfLinkAnnotation>? links)
            ? links
            : Array.Empty<PdfLinkAnnotation>();
    }

    /// <summary>Returns simple named-destination link annotations for a destination name.</summary>
    public IReadOnlyList<PdfLinkAnnotation> GetLinkAnnotationsByDestinationName(string destinationName) {
        Guard.NotNullOrWhiteSpace(destinationName, nameof(destinationName));
        return LinkAnnotationsByDestinationName.TryGetValue(destinationName, out IReadOnlyList<PdfLinkAnnotation>? links)
            ? links
            : Array.Empty<PdfLinkAnnotation>();
    }

    /// <summary>Returns simple direct-destination link annotations for a one-based destination page number.</summary>
    public IReadOnlyList<PdfLinkAnnotation> GetLinkAnnotationsByDestinationPageNumber(int pageNumber) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        return LinkAnnotationsByDestinationPageNumber.TryGetValue(pageNumber, out IReadOnlyList<PdfLinkAnnotation>? links)
            ? links
            : Array.Empty<PdfLinkAnnotation>();
    }

    /// <summary>Returns simple named-action link annotations for a viewer action name.</summary>
    public IReadOnlyList<PdfLinkAnnotation> GetLinkAnnotationsByNamedAction(string namedAction) {
        Guard.NotNullOrWhiteSpace(namedAction, nameof(namedAction));
        return LinkAnnotationsByNamedAction.TryGetValue(namedAction, out IReadOnlyList<PdfLinkAnnotation>? links)
            ? links
            : Array.Empty<PdfLinkAnnotation>();
    }

    /// <summary>Returns simple remote GoTo link annotations for a target file.</summary>
    public IReadOnlyList<PdfLinkAnnotation> GetLinkAnnotationsByRemoteFile(string remoteFile) {
        Guard.NotNullOrWhiteSpace(remoteFile, nameof(remoteFile));
        return LinkAnnotationsByRemoteFile.TryGetValue(remoteFile, out IReadOnlyList<PdfLinkAnnotation>? links)
            ? links
            : Array.Empty<PdfLinkAnnotation>();
    }

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

    /// <summary>True when at least one simple AcroForm widget annotation was read from the document catalog fields.</summary>
    public bool HasFormWidgets => FormWidgetCount > 0;

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

    /// <summary>Returns simple widget annotations for a fully qualified form field name.</summary>
    public IReadOnlyList<PdfFormWidget> GetFormWidgets(string fieldName) {
        Guard.NotNullOrWhiteSpace(fieldName, nameof(fieldName));
        return FormWidgetsByFieldName.TryGetValue(fieldName, out IReadOnlyList<PdfFormWidget>? widgets)
            ? widgets
            : Array.Empty<PdfFormWidget>();
    }

    /// <summary>Returns simple widget annotations for a one-based page number.</summary>
    public IReadOnlyList<PdfFormWidget> GetFormWidgets(int pageNumber) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        return FormWidgetsByPageNumber.TryGetValue(pageNumber, out IReadOnlyList<PdfFormWidget>? widgets)
            ? widgets
            : Array.Empty<PdfFormWidget>();
    }

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

    /// <summary>Effective PDF version inferred from the highest catalog override or file header version.</summary>
    public string? EffectiveVersion => ComparePdfVersion(CatalogVersion, HeaderVersion) >= 0 ? CatalogVersion : HeaderVersion;

    /// <summary>True when the effective PDF version is PDF 2.0 or later.</summary>
    public bool IsPdf20OrLater => ComparePdfVersion(EffectiveVersion, "2.0") >= 0;

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

    private bool HasAcroFormSignatureFlag(int flag) {
        return AcroFormSignatureFlags.HasValue && (AcroFormSignatureFlags.Value & flag) != 0;
    }

    private static int ComparePdfVersion(string? left, string? right) {
        if (!TryParsePdfVersion(left, out int leftMajor, out int leftMinor)) {
            return TryParsePdfVersion(right, out _, out _) ? -1 : 0;
        }

        if (!TryParsePdfVersion(right, out int rightMajor, out int rightMinor)) {
            return 1;
        }

        int majorComparison = leftMajor.CompareTo(rightMajor);
        return majorComparison != 0 ? majorComparison : leftMinor.CompareTo(rightMinor);
    }

    private static bool TryParsePdfVersion(string? version, out int major, out int minor) {
        major = 0;
        minor = 0;
        if (string.IsNullOrWhiteSpace(version)) {
            return false;
        }

        string[] parts = version!.Split('.');
        return parts.Length == 2 &&
            int.TryParse(parts[0], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out major) &&
            int.TryParse(parts[1], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out minor);
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

    private static void AddAnnotation(Dictionary<string, List<PdfAnnotation>> grouped, string actionType, PdfAnnotation annotation) {
        if (!grouped.TryGetValue(actionType, out List<PdfAnnotation>? annotations)) {
            annotations = new List<PdfAnnotation>();
            grouped.Add(actionType, annotations);
        }

        if (!annotations.Contains(annotation)) {
            annotations.Add(annotation);
        }
    }

    private static System.Collections.ObjectModel.ReadOnlyDictionary<TKey, IReadOnlyList<TValue>> ToReadOnlyLookup<TKey, TValue>(Dictionary<TKey, List<TValue>> grouped) where TKey : notnull {
        var result = new Dictionary<TKey, IReadOnlyList<TValue>>(grouped.Count, grouped.Comparer);
        foreach (var item in grouped) {
            result.Add(item.Key, item.Value.AsReadOnly());
        }

        return new System.Collections.ObjectModel.ReadOnlyDictionary<TKey, IReadOnlyList<TValue>>(result);
    }
}

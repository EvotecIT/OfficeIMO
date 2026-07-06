namespace OfficeIMO.Pdf;

/// <summary>
/// Configures reusable proof checks for source-document to PDF conversion results.
/// </summary>
public sealed class PdfConversionProofOptions {
    private readonly List<string> _requiredTextMarkers = new List<string>();
    private readonly List<string> _requiredLogicalSignals = new List<string>();
    private readonly List<string> _requiredOutlineTitles = new List<string>();
    private readonly List<string> _requiredLinkUris = new List<string>();
    private readonly List<string> _requiredFormFieldNames = new List<string>();
    private readonly List<string> _requiredNamedDestinationNames = new List<string>();
    private readonly List<PdfPageLabelRange> _requiredPageLabelRanges = new List<PdfPageLabelRange>();
    private readonly List<string> _requiredAttachmentFileNames = new List<string>();
    private readonly List<string> _requiredOutputIntentSubtypes = new List<string>();
    private readonly List<string> _requiredOutputConditionIdentifiers = new List<string>();
    private readonly List<string> _requiredOptionalContentGroupNames = new List<string>();
    private readonly List<string> _requiredOptionalContentVisibleGroupNames = new List<string>();
    private readonly List<string> _requiredOptionalContentHiddenGroupNames = new List<string>();
    private readonly List<string> _requiredOptionalContentLockedGroupNames = new List<string>();
    private readonly List<string> _requiredOptionalContentOrderedGroupNames = new List<string>();
    private readonly List<string> _requiredXmpSubjects = new List<string>();
    private readonly List<string> _requiredTaggedStructureTypes = new List<string>();
    private readonly Dictionary<string, string> _requiredViewerPreferences = new Dictionary<string, string>(StringComparer.Ordinal);
    private readonly List<string> _requiredWarningCodes = new List<string>();
    private readonly List<string> _requiredWarningSources = new List<string>();
    private readonly HashSet<string> _acceptedWarningCodes = new HashSet<string>(StringComparer.Ordinal);

    /// <summary>True when the generated PDF must be readable by the OfficeIMO.Pdf reader.</summary>
    public bool RequireReadablePdf { get; set; } = true;

    /// <summary>True when error-severity conversion warnings should fail the proof.</summary>
    public bool RequireNoErrorWarnings { get; set; }

    /// <summary>True when conversion warnings whose codes are not accepted should fail the proof.</summary>
    public bool RequireNoUnexpectedWarnings { get; set; }

    /// <summary>True when the proof snapshot should include the generated artifact byte count and SHA-256.</summary>
    public bool IncludeArtifactHash { get; set; } = true;

    /// <summary>Expected SHA-256 hash for the generated artifact.</summary>
    public string? RequiredArtifactSha256 { get; set; }

    /// <summary>Expected page count for the generated PDF artifact.</summary>
    public int? RequiredPageCount { get; set; }

    /// <summary>Expected page width in PDF points for every generated page.</summary>
    public double? RequiredPageWidth { get; set; }

    /// <summary>Expected page height in PDF points for every generated page.</summary>
    public double? RequiredPageHeight { get; set; }

    /// <summary>Allowed point difference when comparing required page size.</summary>
    public double RequiredPageSizeTolerance { get; set; } = 0.01D;

    /// <summary>Expected PDF Info dictionary title.</summary>
    public string? RequiredMetadataTitle { get; set; }

    /// <summary>Expected PDF Info dictionary author.</summary>
    public string? RequiredMetadataAuthor { get; set; }

    /// <summary>Expected PDF Info dictionary subject.</summary>
    public string? RequiredMetadataSubject { get; set; }

    /// <summary>Expected PDF Info dictionary keywords.</summary>
    public string? RequiredMetadataKeywords { get; set; }

    /// <summary>Expected catalog language value.</summary>
    public string? RequiredCatalogLanguage { get; set; }

    /// <summary>Expected catalog page mode PDF name.</summary>
    public string? RequiredCatalogPageMode { get; set; }

    /// <summary>Expected catalog page layout PDF name.</summary>
    public string? RequiredCatalogPageLayout { get; set; }

    /// <summary>Expected one-based document open-action target page number.</summary>
    public int? RequiredOpenActionPageNumber { get; set; }

    /// <summary>Expected document open-action destination mode.</summary>
    public PdfOpenActionDestinationMode? RequiredOpenActionDestinationMode { get; set; }

    /// <summary>Expected optional-content default configuration name.</summary>
    public string? RequiredOptionalContentDefaultConfigurationName { get; set; }

    /// <summary>Expected optional-content default configuration creator.</summary>
    public string? RequiredOptionalContentDefaultConfigurationCreator { get; set; }

    /// <summary>Expected optional-content default configuration base state.</summary>
    public string? RequiredOptionalContentBaseState { get; set; }

    /// <summary>Minimum optional-content group count required in the generated PDF.</summary>
    public int? RequiredOptionalContentGroupCountAtLeast { get; set; }

    /// <summary>Expected XMP Dublin Core title.</summary>
    public string? RequiredXmpTitle { get; set; }

    /// <summary>Expected XMP Dublin Core creator.</summary>
    public string? RequiredXmpCreator { get; set; }

    /// <summary>Expected XMP Dublin Core description.</summary>
    public string? RequiredXmpDescription { get; set; }

    /// <summary>Expected XMP PDF producer metadata.</summary>
    public string? RequiredXmpProducer { get; set; }

    /// <summary>Expected XMP PDF keywords metadata.</summary>
    public string? RequiredXmpKeywords { get; set; }

    /// <summary>Expected XMP PDF/A identification part.</summary>
    public int? RequiredXmpPdfAPart { get; set; }

    /// <summary>Expected XMP PDF/A identification conformance.</summary>
    public string? RequiredXmpPdfAConformance { get; set; }

    /// <summary>Expected XMP PDF/UA identification part.</summary>
    public int? RequiredXmpPdfUaPart { get; set; }

    /// <summary>Minimum tagged-PDF structure element count required in the generated PDF.</summary>
    public int? RequiredTaggedStructureElementCountAtLeast { get; set; }

    /// <summary>Minimum tagged-PDF marked-content reference count required in the generated PDF.</summary>
    public int? RequiredTaggedMarkedContentReferencesAtLeast { get; set; }

    /// <summary>Text markers that must be extractable from the generated PDF.</summary>
    public IList<string> RequiredTextMarkers => _requiredTextMarkers;

    /// <summary>Logical readback signals that must be present in the generated PDF.</summary>
    public IList<string> RequiredLogicalSignals => _requiredLogicalSignals;

    /// <summary>PDF outline/bookmark titles that must be present in the generated PDF.</summary>
    public IList<string> RequiredOutlineTitles => _requiredOutlineTitles;

    /// <summary>URI link targets that must be present in the generated PDF.</summary>
    public IList<string> RequiredLinkUris => _requiredLinkUris;

    /// <summary>AcroForm field names that must be present in the generated PDF.</summary>
    public IList<string> RequiredFormFieldNames => _requiredFormFieldNames;

    /// <summary>Named destinations that must be present in the generated PDF catalog.</summary>
    public IList<string> RequiredNamedDestinationNames => _requiredNamedDestinationNames;

    /// <summary>Page-label ranges that must be present in the generated PDF catalog.</summary>
    public IList<PdfPageLabelRange> RequiredPageLabelRanges => _requiredPageLabelRanges;

    /// <summary>Embedded or associated file names that must be present in the generated PDF catalog.</summary>
    public IList<string> RequiredAttachmentFileNames => _requiredAttachmentFileNames;

    /// <summary>Output intent subtypes that must be present in the generated PDF catalog.</summary>
    public IList<string> RequiredOutputIntentSubtypes => _requiredOutputIntentSubtypes;

    /// <summary>Output condition identifiers that must be present in the generated PDF catalog.</summary>
    public IList<string> RequiredOutputConditionIdentifiers => _requiredOutputConditionIdentifiers;

    /// <summary>Optional-content/layer group names that must be present in the generated PDF catalog.</summary>
    public IList<string> RequiredOptionalContentGroupNames => _requiredOptionalContentGroupNames;

    /// <summary>Optional-content/layer group names that must initially be visible.</summary>
    public IList<string> RequiredOptionalContentVisibleGroupNames => _requiredOptionalContentVisibleGroupNames;

    /// <summary>Optional-content/layer group names that must initially be hidden.</summary>
    public IList<string> RequiredOptionalContentHiddenGroupNames => _requiredOptionalContentHiddenGroupNames;

    /// <summary>Optional-content/layer group names that must be locked in the default configuration.</summary>
    public IList<string> RequiredOptionalContentLockedGroupNames => _requiredOptionalContentLockedGroupNames;

    /// <summary>Optional-content/layer group names that must appear in the default configuration order.</summary>
    public IList<string> RequiredOptionalContentOrderedGroupNames => _requiredOptionalContentOrderedGroupNames;

    /// <summary>XMP Dublin Core subjects that must be present in the generated PDF catalog metadata.</summary>
    public IList<string> RequiredXmpSubjects => _requiredXmpSubjects;

    /// <summary>Tagged-PDF structure types that must be present in the generated PDF structure tree.</summary>
    public IList<string> RequiredTaggedStructureTypes => _requiredTaggedStructureTypes;

    /// <summary>Viewer preference values that must be present in the generated PDF catalog.</summary>
    public IDictionary<string, string> RequiredViewerPreferences => _requiredViewerPreferences;

    /// <summary>Stable warning codes that must be present in the captured conversion report.</summary>
    public IList<string> RequiredWarningCodes => _requiredWarningCodes;

    /// <summary>Warning source labels that must be present in the captured conversion report.</summary>
    public IList<string> RequiredWarningSources => _requiredWarningSources;

    /// <summary>Stable warning codes accepted by this proof contract when unexpected warnings are rejected.</summary>
    public ISet<string> AcceptedWarningCodes => _acceptedWarningCodes;

    /// <summary>Adds required text markers and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireTextMarkers(params string[] markers) {
        AddValues(_requiredTextMarkers, markers);
        return this;
    }

    /// <summary>Adds required logical readback signals and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireLogicalSignals(params string[] signals) {
        AddValues(_requiredLogicalSignals, signals);
        return this;
    }

    /// <summary>Adds required outline/bookmark titles and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireOutlineTitles(params string[] titles) {
        AddValues(_requiredOutlineTitles, titles);
        return this;
    }

    /// <summary>Adds required URI link targets and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireLinkUris(params string[] uris) {
        AddValues(_requiredLinkUris, uris);
        return this;
    }

    /// <summary>Adds required AcroForm field names and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireFormFieldNames(params string[] names) {
        AddValues(_requiredFormFieldNames, names);
        return this;
    }

    /// <summary>Adds required named destinations and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireNamedDestinationNames(params string[] names) {
        AddValues(_requiredNamedDestinationNames, names);
        return this;
    }

    /// <summary>Adds a required page-label range and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequirePageLabelRange(int startPageNumber, PdfPageNumberStyle style, int startNumber = 1, string? prefix = null) {
        _requiredPageLabelRanges.Add(new PdfPageLabelRange(startPageNumber, style, startNumber, prefix));
        return this;
    }

    /// <summary>Adds required page-label ranges and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequirePageLabelRanges(params PdfPageLabelRange[] ranges) {
        Guard.NotNull(ranges, nameof(ranges));
        for (int i = 0; i < ranges.Length; i++) {
            if (ranges[i] is not null) {
                _requiredPageLabelRanges.Add(ranges[i]);
            }
        }

        return this;
    }

    /// <summary>Adds required embedded or associated file names and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireAttachmentFileNames(params string[] fileNames) {
        AddValues(_requiredAttachmentFileNames, fileNames);
        return this;
    }

    /// <summary>Adds required output intent subtypes and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireOutputIntentSubtypes(params string[] subtypes) {
        AddValues(_requiredOutputIntentSubtypes, subtypes);
        return this;
    }

    /// <summary>Adds required output condition identifiers and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireOutputConditionIdentifiers(params string[] identifiers) {
        AddValues(_requiredOutputConditionIdentifiers, identifiers);
        return this;
    }

    /// <summary>Adds required optional-content/layer group names and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireOptionalContentGroupNames(params string[] names) {
        AddValues(_requiredOptionalContentGroupNames, names);
        return this;
    }

    /// <summary>Requires at least the supplied number of optional-content/layer groups.</summary>
    public PdfConversionProofOptions RequireOptionalContentGroupCountAtLeast(int count) {
        if (count < 1) {
            throw new ArgumentOutOfRangeException(nameof(count), count, "PDF conversion proof optional-content group count must be at least 1.");
        }

        RequiredOptionalContentGroupCountAtLeast = count;
        return this;
    }

    /// <summary>Requires selected optional-content default configuration values to match the generated PDF.</summary>
    public PdfConversionProofOptions RequireOptionalContentDefaultConfiguration(string? name = null, string? creator = null, string? baseState = null) {
        if (name is not null) {
            RequiredOptionalContentDefaultConfigurationName = name;
        }

        if (creator is not null) {
            RequiredOptionalContentDefaultConfigurationCreator = creator;
        }

        if (baseState is not null) {
            RequiredOptionalContentBaseState = baseState;
        }

        return this;
    }

    /// <summary>Adds optional-content/layer group names that must initially be visible.</summary>
    public PdfConversionProofOptions RequireOptionalContentVisibleGroupNames(params string[] names) {
        AddValues(_requiredOptionalContentVisibleGroupNames, names);
        return this;
    }

    /// <summary>Adds optional-content/layer group names that must initially be hidden.</summary>
    public PdfConversionProofOptions RequireOptionalContentHiddenGroupNames(params string[] names) {
        AddValues(_requiredOptionalContentHiddenGroupNames, names);
        return this;
    }

    /// <summary>Adds optional-content/layer group names that must be locked in the default configuration.</summary>
    public PdfConversionProofOptions RequireOptionalContentLockedGroupNames(params string[] names) {
        AddValues(_requiredOptionalContentLockedGroupNames, names);
        return this;
    }

    /// <summary>Adds optional-content/layer group names that must appear in the default configuration order.</summary>
    public PdfConversionProofOptions RequireOptionalContentOrderedGroupNames(params string[] names) {
        AddValues(_requiredOptionalContentOrderedGroupNames, names);
        return this;
    }

    /// <summary>Requires selected XMP metadata values to match the generated PDF catalog metadata stream.</summary>
    public PdfConversionProofOptions RequireXmpMetadata(string? title = null, string? creator = null, string? description = null, string? producer = null, string? keywords = null) {
        if (title is not null) {
            RequiredXmpTitle = title;
        }

        if (creator is not null) {
            RequiredXmpCreator = creator;
        }

        if (description is not null) {
            RequiredXmpDescription = description;
        }

        if (producer is not null) {
            RequiredXmpProducer = producer;
        }

        if (keywords is not null) {
            RequiredXmpKeywords = keywords;
        }

        return this;
    }

    /// <summary>Adds required XMP subject values and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireXmpSubjects(params string[] subjects) {
        AddValues(_requiredXmpSubjects, subjects);
        return this;
    }

    /// <summary>Requires XMP PDF/A identification metadata to match the generated PDF.</summary>
    public PdfConversionProofOptions RequireXmpPdfAIdentification(int part, string? conformance = null) {
        if (part < 1) {
            throw new ArgumentOutOfRangeException(nameof(part), part, "PDF conversion proof PDF/A part must be at least 1.");
        }

        RequiredXmpPdfAPart = part;
        if (conformance is not null) {
            RequiredXmpPdfAConformance = conformance;
        }

        return this;
    }

    /// <summary>Requires XMP PDF/UA identification metadata to match the generated PDF.</summary>
    public PdfConversionProofOptions RequireXmpPdfUaIdentification(int part = 1) {
        if (part < 1) {
            throw new ArgumentOutOfRangeException(nameof(part), part, "PDF conversion proof PDF/UA part must be at least 1.");
        }

        RequiredXmpPdfUaPart = part;
        return this;
    }

    /// <summary>Adds required tagged-PDF structure types and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireTaggedStructureTypes(params string[] structureTypes) {
        AddValues(_requiredTaggedStructureTypes, structureTypes);
        return this;
    }

    /// <summary>Requires at least the supplied number of tagged-PDF structure elements.</summary>
    public PdfConversionProofOptions RequireTaggedStructureElementCountAtLeast(int count) {
        if (count < 1) {
            throw new ArgumentOutOfRangeException(nameof(count), count, "PDF conversion proof tagged structure element count must be at least 1.");
        }

        RequiredTaggedStructureElementCountAtLeast = count;
        return this;
    }

    /// <summary>Requires at least the supplied number of tagged-PDF marked-content references.</summary>
    public PdfConversionProofOptions RequireTaggedMarkedContentReferencesAtLeast(int count) {
        if (count < 1) {
            throw new ArgumentOutOfRangeException(nameof(count), count, "PDF conversion proof tagged marked-content reference count must be at least 1.");
        }

        RequiredTaggedMarkedContentReferencesAtLeast = count;
        return this;
    }

    /// <summary>Requires the generated PDF catalog language to match the supplied value.</summary>
    public PdfConversionProofOptions RequireCatalogLanguage(string language) {
        RequiredCatalogLanguage = string.IsNullOrWhiteSpace(language) ? string.Empty : language.Trim();
        return this;
    }

    /// <summary>Requires the generated PDF catalog view hints to match the supplied values.</summary>
    public PdfConversionProofOptions RequireCatalogView(PdfCatalogPageMode? pageMode = null, PdfCatalogPageLayout? pageLayout = null) {
        if (pageMode.HasValue) {
            RequiredCatalogPageMode = PdfCatalogDictionaryBuilder.GetPageModeName(pageMode.Value);
        }

        if (pageLayout.HasValue) {
            RequiredCatalogPageLayout = PdfCatalogDictionaryBuilder.GetPageLayoutName(pageLayout.Value);
        }

        return this;
    }

    /// <summary>Requires the generated PDF catalog page mode to match the supplied value.</summary>
    public PdfConversionProofOptions RequireCatalogPageMode(PdfCatalogPageMode pageMode) {
        RequiredCatalogPageMode = PdfCatalogDictionaryBuilder.GetPageModeName(pageMode);
        return this;
    }

    /// <summary>Requires the generated PDF catalog page layout to match the supplied value.</summary>
    public PdfConversionProofOptions RequireCatalogPageLayout(PdfCatalogPageLayout pageLayout) {
        RequiredCatalogPageLayout = PdfCatalogDictionaryBuilder.GetPageLayoutName(pageLayout);
        return this;
    }

    /// <summary>Requires the generated PDF catalog open action to target the supplied page and optional destination mode.</summary>
    public PdfConversionProofOptions RequireOpenAction(int pageNumber, PdfOpenActionDestinationMode? destinationMode = null) {
        if (pageNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "PDF conversion proof open-action page number must be at least 1.");
        }

        if (destinationMode.HasValue) {
            ValidateOpenActionDestinationMode(destinationMode.Value, nameof(destinationMode));
        }

        RequiredOpenActionPageNumber = pageNumber;
        RequiredOpenActionDestinationMode = destinationMode;
        return this;
    }

    /// <summary>Requires a generated PDF catalog viewer preference value.</summary>
    public PdfConversionProofOptions RequireViewerPreference(string name, string value) {
        if (!string.IsNullOrWhiteSpace(name)) {
            _requiredViewerPreferences[name] = value ?? string.Empty;
        }

        return this;
    }

    /// <summary>Requires a generated PDF catalog viewer preference boolean value.</summary>
    public PdfConversionProofOptions RequireViewerPreference(string name, bool value) {
        return RequireViewerPreference(name, value ? "true" : "false");
    }

    /// <summary>Adds required warning codes and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireWarningCodes(params string[] codes) {
        AddValues(_requiredWarningCodes, codes);
        return this;
    }

    /// <summary>Adds required warning sources and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions RequireWarningSources(params string[] sources) {
        AddValues(_requiredWarningSources, sources);
        return this;
    }

    /// <summary>Adds accepted warning codes and returns this options object for fluent setup.</summary>
    public PdfConversionProofOptions AcceptWarningCodes(params string[] codes) {
        AddValues(_acceptedWarningCodes, codes);
        return this;
    }

    /// <summary>Requires every conversion warning code to be explicitly accepted.</summary>
    public PdfConversionProofOptions RequireNoUnexpectedWarningCodes() {
        RequireNoUnexpectedWarnings = true;
        return this;
    }

    /// <summary>Requires the generated artifact SHA-256 to match the supplied hash.</summary>
    public PdfConversionProofOptions RequireArtifactSha256(string sha256) {
        RequiredArtifactSha256 = string.IsNullOrWhiteSpace(sha256) ? string.Empty : sha256.Trim().ToLowerInvariant();
        return this;
    }

    /// <summary>Requires the generated PDF to contain exactly the supplied number of pages.</summary>
    public PdfConversionProofOptions RequirePageCount(int pageCount) {
        if (pageCount < 1) {
            throw new ArgumentOutOfRangeException(nameof(pageCount), pageCount, "PDF conversion proof page count must be at least 1.");
        }

        RequiredPageCount = pageCount;
        return this;
    }

    /// <summary>Requires every generated PDF page to match the supplied size in PDF points.</summary>
    public PdfConversionProofOptions RequirePageSize(double width, double height, double tolerance = 0.01D) {
        ValidatePositiveFinite(width, nameof(width), "PDF conversion proof page width must be a positive finite value.");
        ValidatePositiveFinite(height, nameof(height), "PDF conversion proof page height must be a positive finite value.");
        ValidateNonNegativeFinite(tolerance, nameof(tolerance), "PDF conversion proof page-size tolerance must be a non-negative finite value.");

        RequiredPageWidth = width;
        RequiredPageHeight = height;
        RequiredPageSizeTolerance = tolerance;
        return this;
    }

    /// <summary>Requires selected PDF Info dictionary metadata values to match the generated PDF.</summary>
    public PdfConversionProofOptions RequireMetadata(string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        if (title is not null) {
            RequiredMetadataTitle = title;
        }

        if (author is not null) {
            RequiredMetadataAuthor = author;
        }

        if (subject is not null) {
            RequiredMetadataSubject = subject;
        }

        if (keywords is not null) {
            RequiredMetadataKeywords = keywords;
        }

        return this;
    }

    /// <summary>Requires the conversion report to contain no error-severity warnings.</summary>
    public PdfConversionProofOptions RequireNoErrors() {
        RequireNoErrorWarnings = true;
        return this;
    }

    private static void AddValues(List<string> target, string[] values) {
        Guard.NotNull(values, nameof(values));
        for (int i = 0; i < values.Length; i++) {
            if (!string.IsNullOrWhiteSpace(values[i])) {
                target.Add(values[i]);
            }
        }
    }

    private static void AddValues(HashSet<string> target, string[] values) {
        Guard.NotNull(values, nameof(values));
        for (int i = 0; i < values.Length; i++) {
            if (!string.IsNullOrWhiteSpace(values[i])) {
                target.Add(values[i]);
            }
        }
    }

    private static void ValidatePositiveFinite(double value, string parameterName, string message) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new ArgumentOutOfRangeException(parameterName, value, message);
        }
    }

    private static void ValidateNonNegativeFinite(double value, string parameterName, string message) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(parameterName, value, message);
        }
    }

    private static void ValidateOpenActionDestinationMode(PdfOpenActionDestinationMode value, string parameterName) {
        if (value < PdfOpenActionDestinationMode.Xyz || value > PdfOpenActionDestinationMode.FitBoundingBoxVertical) {
            throw new ArgumentOutOfRangeException(parameterName, value, "PDF conversion proof open-action destination mode is not supported.");
        }
    }
}

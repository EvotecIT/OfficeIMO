namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>When true, generated page content streams are written with Flate compression.</summary>
    public bool CompressContentStreams { get; set; }
    private long _objectBufferMemoryLimitBytes = PdfObjectStore.DefaultMemoryLimitBytes;
    private long _pageContentMemoryLimitBytes = PdfPageContentStore.DefaultMemoryLimitBytes;
    private PdfObjectSerializationMode _objectSerializationMode;
    /// <summary>Controls whether completed indirect objects are buffered or emitted once in forward-only order.</summary>
    public PdfObjectSerializationMode ObjectSerializationMode {
        get => _objectSerializationMode;
        set {
            if (value != PdfObjectSerializationMode.Buffered &&
                value != PdfObjectSerializationMode.ForwardOnly) {
                throw new System.ArgumentOutOfRangeException(nameof(ObjectSerializationMode), value, "PDF object serialization mode must be Buffered or ForwardOnly.");
            }
            _objectSerializationMode = value;
        }
    }
    /// <summary>Maximum serialized-object bytes retained in memory while saving. Zero spills every completed object to temporary storage.</summary>
    public long ObjectBufferMemoryLimitBytes {
        get => _objectBufferMemoryLimitBytes;
        set {
            if (value < 0L) throw new System.ArgumentOutOfRangeException(nameof(ObjectBufferMemoryLimitBytes), value, "PDF object-buffer memory limit cannot be negative.");
            _objectBufferMemoryLimitBytes = value;
        }
    }
    /// <summary>Maximum completed page-content bytes retained during layout. Zero spills every completed page and effect stream to temporary storage.</summary>
    public long PageContentMemoryLimitBytes {
        get => _pageContentMemoryLimitBytes;
        set {
            if (value < 0L) throw new System.ArgumentOutOfRangeException(nameof(PageContentMemoryLimitBytes), value, "PDF page-content memory limit cannot be negative.");
            _pageContentMemoryLimitBytes = value;
        }
    }
    /// <summary>When true, generated standard-font resources include WinAnsi-to-Unicode CMaps for stronger extraction and compliance groundwork.</summary>
    public bool IncludeStandardFontToUnicodeMaps { get; set; }
    /// <summary>When true, embedded TrueType font file streams are Flate-compressed while preserving their original /Length1 metadata.</summary>
    public bool CompressEmbeddedFonts { get; set; } = true;
    /// <summary>When true, generated PDFs include a catalog XMP metadata stream synchronized with document Info metadata.</summary>
    public bool IncludeXmpMetadata { get; set; }
    /// <summary>When true, generated PDFs include catalog page labels that match the configured page-number style and start number.</summary>
    public bool IncludePageLabels { get; set; }
    /// <summary>When true, generated FreeText and Highlight annotations are painted into page content instead of emitted as interactive annotations.</summary>
    public bool FlattenVisualAnnotations { get; set; }
    /// <summary>Optional catalog page-label prefix, for example "A-" or "Appendix ". Requires <see cref="IncludePageLabels"/> to be emitted.</summary>
    public string? PageLabelPrefix {
        get => _pageLabelPrefix;
        set {
            PdfPageLabelDictionaryBuilder.ValidatePrefix(value, nameof(PageLabelPrefix));
            _pageLabelPrefix = value;
        }
    }

    /// <summary>Generated catalog page-label rules, ordered by their one-based start page number.</summary>
    public System.Collections.Generic.IReadOnlyList<PdfPageLabelRange> PageLabelRanges =>
        _pageLabelRanges == null || _pageLabelRanges.Count == 0
            ? System.Array.Empty<PdfPageLabelRange>()
            : _pageLabelRanges.OrderBy(range => range.StartPageNumber).ToList().AsReadOnly();

    internal System.Collections.Generic.IReadOnlyList<PdfPageLabelRange> PageLabelRangeSnapshots =>
        _pageLabelRanges == null || _pageLabelRanges.Count == 0
            ? System.Array.Empty<PdfPageLabelRange>()
            : _pageLabelRanges.OrderBy(range => range.StartPageNumber).ToList().AsReadOnly();
    /// <summary>PDF file header version emitted for generated documents.</summary>
    public PdfFileVersion FileVersion {
        get => _fileVersion;
        set {
            Guard.FileVersion(value, nameof(FileVersion));
            _fileVersion = value;
        }
    }
    /// <summary>Requested generated-PDF compliance profile. Non-None profiles are validated strictly and fail until their required primitives are implemented.</summary>
    public PdfComplianceProfile ComplianceProfile {
        get => _complianceProfile;
        set {
            Guard.ComplianceProfile(value, nameof(ComplianceProfile));
            _complianceProfile = value;
        }
    }
    /// <summary>Optional PDF/A XMP identification metadata. This does not by itself certify PDF/A conformance.</summary>
    public PdfAIdentification? PdfAIdentification {
        get => _pdfAIdentification?.Clone();
        set => _pdfAIdentification = value?.Clone();
    }
    internal PdfAIdentification? PdfAIdentificationSnapshot => _pdfAIdentification?.Clone();
    /// <summary>Optional PDF/UA XMP identification metadata. This does not by itself certify PDF/UA conformance.</summary>
    public PdfUaIdentification? PdfUaIdentification {
        get => _pdfUaIdentification?.Clone();
        set => _pdfUaIdentification = value?.Clone();
    }
    internal PdfUaIdentification? PdfUaIdentificationSnapshot => _pdfUaIdentification?.Clone();
    /// <summary>Optional Factur-X/ZUGFeRD XMP extension metadata. This does not by itself certify e-invoice conformance.</summary>
    public PdfElectronicInvoiceMetadata? ElectronicInvoiceMetadata {
        get => _electronicInvoiceMetadata?.Clone();
        set => _electronicInvoiceMetadata = value?.Clone();
    }
    internal PdfElectronicInvoiceMetadata? ElectronicInvoiceMetadataSnapshot => _electronicInvoiceMetadata?.Clone();
    /// <summary>Optional generated catalog output intent backed by an ICC profile.</summary>
    public PdfOutputIntent? OutputIntent {
        get => _outputIntent?.Clone();
        set => _outputIntent = value?.Clone();
    }
    internal PdfOutputIntent? OutputIntentSnapshot => _outputIntent?.Clone();
    /// <summary>Controls catalog-level tagged-PDF groundwork emitted for accessibility-oriented profiles.</summary>
    public PdfTaggedStructureMode TaggedStructureMode {
        get => _taggedStructureMode;
        set {
            Guard.TaggedStructureMode(value, nameof(TaggedStructureMode));
            _taggedStructureMode = value;
        }
    }
    /// <summary>Optional document language for the generated catalog /Lang entry, for example "en-US".</summary>
    public string? Language {
        get => _language;
        set {
            ValidateOptionalLanguage(value, nameof(Language));
            _language = value;
        }
    }

    /// <summary>Optional catalog page mode emitted for generated PDFs.</summary>
    public PdfCatalogPageMode? CatalogPageMode {
        get => _catalogPageMode;
        set {
            if (value.HasValue) {
                Guard.CatalogPageMode(value.Value, nameof(CatalogPageMode));
            }

            _catalogPageMode = value;
        }
    }

    internal PdfCatalogPageMode? CatalogPageModeSnapshot => _catalogPageMode;

    /// <summary>Optional catalog page layout emitted for generated PDFs.</summary>
    public PdfCatalogPageLayout? CatalogPageLayout {
        get => _catalogPageLayout;
        set {
            if (value.HasValue) {
                Guard.CatalogPageLayout(value.Value, nameof(CatalogPageLayout));
            }

            _catalogPageLayout = value;
        }
    }

    internal PdfCatalogPageLayout? CatalogPageLayoutSnapshot => _catalogPageLayout;

    /// <summary>Optional generated catalog open action controlling the initial page destination.</summary>
    public PdfOpenActionOptions? OpenAction {
        get => _openAction?.Clone();
        set => _openAction = value?.Clone();
    }

    internal PdfOpenActionOptions? OpenActionSnapshot => _openAction?.Clone();

    /// <summary>Optional simple viewer preferences emitted through the generated catalog.</summary>
    public PdfViewerPreferencesOptions? ViewerPreferences {
        get => _viewerPreferences?.Clone();
        set => _viewerPreferences = value?.Clone();
    }

    internal PdfViewerPreferencesOptions? ViewerPreferencesSnapshot => _viewerPreferences?.Clone();

    /// <summary>Optional catalog URI base used by viewers to resolve relative URI actions.</summary>
    public string? CatalogUriBase {
        get => _catalogUriBase;
        set {
            ValidateOptionalCatalogUriBase(value, nameof(CatalogUriBase));
            _catalogUriBase = value;
        }
    }

    internal string? CatalogUriBaseSnapshot => _catalogUriBase;

    /// <summary>Optional Standard password security for generated PDFs.</summary>
    public PdfStandardEncryptionOptions? Encryption {
        get => _encryption?.Clone();
        set => _encryption = value?.Clone();
    }

    internal PdfStandardEncryptionOptions? EncryptionSnapshot => _encryption?.Clone();

    /// <summary>Sets Standard password security for generated PDFs.</summary>
    public PdfOptions SetEncryption(string userPassword, string? ownerPassword = null, int permissions = PdfStandardEncryptionOptions.AllowAllPermissions) {
        Encryption = new PdfStandardEncryptionOptions(userPassword) {
            OwnerPassword = ownerPassword,
            Permissions = permissions
        };
        return this;
    }

    /// <summary>Sets or clears Standard password security for generated PDFs.</summary>
    public PdfOptions SetEncryption(PdfStandardEncryptionOptions? encryption) {
        Encryption = encryption;
        return this;
    }

    /// <summary>Clears generated PDF password security.</summary>
    public PdfOptions ClearEncryption() {
        _encryption = null;
        return this;
    }

    /// <summary>Optional AcroForm default text alignment emitted as catalog-level /Q quadding.</summary>
    public PdfFormFieldTextAlignment? AcroFormDefaultTextAlignment {
        get => _acroFormDefaultTextAlignment;
        set {
            if (value.HasValue) {
                Guard.FormFieldTextAlignment(value.Value, nameof(AcroFormDefaultTextAlignment));
            }

            _acroFormDefaultTextAlignment = value;
        }
    }

    internal PdfFormFieldTextAlignment? AcroFormDefaultTextAlignmentSnapshot => _acroFormDefaultTextAlignment;

    /// <summary>Embedded files associated with the generated document catalog.</summary>
    public System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile> EmbeddedFiles =>
        (System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile>?)CloneEmbeddedFiles(_embeddedFiles) ?? System.Array.Empty<PdfEmbeddedFile>();

    internal System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile> EmbeddedFileSnapshots =>
        (System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile>?)CloneEmbeddedFiles(_embeddedFiles) ?? System.Array.Empty<PdfEmbeddedFile>();

    /// <summary>Optional document portfolio configuration for the generated embedded files.</summary>
    public PdfPortfolioOptions? Portfolio {
        get => _portfolio?.Clone();
        set => _portfolio = value?.Clone();
    }

    internal PdfPortfolioOptions? PortfolioSnapshot => _portfolio?.Clone();

    /// <summary>Configures the generated embedded files as a document portfolio.</summary>
    public PdfOptions SetPortfolio(PdfPortfolioOptions portfolio) {
        Guard.NotNull(portfolio, nameof(portfolio));
        Portfolio = portfolio;
        return this;
    }

    /// <summary>Clears the generated document portfolio configuration without removing embedded files.</summary>
    public PdfOptions ClearPortfolio() {
        _portfolio = null;
        return this;
    }

    /// <summary>Embedded TrueType font mappings keyed by generated standard-font slot.</summary>
    public System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, PdfEmbeddedFont> EmbeddedFonts {
        get {
            if (_embeddedFonts == null || _embeddedFonts.Count == 0) {
                return new System.Collections.Generic.Dictionary<PdfStandardFont, PdfEmbeddedFont>();
            }

            var copy = new System.Collections.Generic.Dictionary<PdfStandardFont, PdfEmbeddedFont>();
            foreach (var embeddedFont in _embeddedFonts) {
                copy[embeddedFont.Key] = embeddedFont.Value.Clone();
            }

            return copy;
        }
    }

    /// <summary>Embeds a TrueType font file for a generated standard-font slot.</summary>
    public PdfOptions EmbedStandardFont(PdfStandardFont font, byte[] data, string? fontName = null) {
        var embeddedFont = new PdfEmbeddedFont(font, data, fontName);
        (_embeddedFonts ??= new System.Collections.Generic.Dictionary<PdfStandardFont, PdfEmbeddedFont>())[font] = embeddedFont;
        _embeddedFontPrograms?.Remove(font);
        _embeddedOpenTypeCffFontPrograms?.Remove(font);
        _embeddedFontProgramFailures?.Remove(font);
        ClearReportedEmbeddedFontProgramFailure(font);
        return this;
    }

    /// <summary>Embeds a TrueType font file from disk for a generated standard-font slot.</summary>
    public PdfOptions EmbedStandardFont(PdfStandardFont font, string path, string? fontName = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return EmbedStandardFont(font, System.IO.File.ReadAllBytes(path), fontName);
    }

    /// <summary>
    /// Uses a caller-supplied TrueType font family as the generated document's default font family.
    /// </summary>
    /// <remarks>
    /// The current implementation maps the family variants onto the writer's generated Helvetica
    /// family slots, so existing paragraph, heading, rich-run, table, header, footer, form, and
    /// watermark layout paths can use the embedded font metrics without requiring callers to know
    /// those internal slots. Missing bold or italic variants fall back to the regular face.
    /// </remarks>
    public PdfOptions UseFontFamily(
        string familyName,
        byte[] regular,
        byte[]? bold = null,
        byte[]? italic = null,
        byte[]? boldItalic = null) {
        return UseFontFamily(new PdfEmbeddedFontFamily(familyName, regular, bold, italic, boldItalic));
    }

    /// <summary>
    /// Uses a reusable caller-supplied TrueType font family as the generated document's default font family.
    /// </summary>
    public PdfOptions UseFontFamily(PdfEmbeddedFontFamily fontFamily) {
        UseDefaultTextFontFamily(fontFamily);
        HeaderFont = PdfStandardFont.Helvetica;
        HeaderFontFamily = null;
        FooterFont = PdfStandardFont.Helvetica;
        FooterFontFamily = null;
        return this;
    }

    /// <summary>
    /// Uses caller-supplied TrueType font files as the generated document's default font family.
    /// </summary>
    public PdfOptions UseFontFamily(
        string familyName,
        string regularPath,
        string? boldPath = null,
        string? italicPath = null,
        string? boldItalicPath = null) {
        return UseFontFamily(PdfEmbeddedFontFamily.FromFiles(familyName, regularPath, boldPath, italicPath, boldItalicPath));
    }

    /// <summary>Removes all embedded standard-font mappings.</summary>
    public PdfOptions ClearEmbeddedStandardFonts() {
        _embeddedFonts?.Clear();
        _embeddedFontPrograms?.Clear();
        _embeddedOpenTypeCffFontPrograms?.Clear();
        _embeddedFontProgramFailures?.Clear();
        _reportedEmbeddedFontProgramFailures?.Clear();
        _embeddedFontFallbacks = null;
        return this;
    }

    internal PdfOptions UseDefaultTextFontFamily(PdfEmbeddedFontFamily fontFamily) {
        RegisterFontFamily(PdfStandardFont.Helvetica, fontFamily);
        DefaultFont = PdfStandardFont.Helvetica;
        return this;
    }

    private static string BuildFontFamilyFaceName(string familyName, string faceName) =>
        familyName + "-" + faceName;

    /// <summary>Requests a generated-PDF compliance profile for this document.</summary>
    public PdfOptions RequireCompliance(PdfComplianceProfile profile) {
        ComplianceProfile = profile;
        return this;
    }

    /// <summary>Sets the PDF file header version emitted for generated documents.</summary>
    public PdfOptions SetFileVersion(PdfFileVersion version) {
        FileVersion = version;
        return this;
    }

    /// <summary>Sets PDF/A XMP identification metadata and enables XMP metadata emission.</summary>
    public PdfOptions SetPdfAIdentification(PdfAIdentification? identification) {
        PdfAIdentification = identification;
        if (identification != null) {
            IncludeXmpMetadata = true;
        }

        return this;
    }

    /// <summary>Sets PDF/A XMP identification metadata and enables XMP metadata emission.</summary>
    public PdfOptions SetPdfAIdentification(int part, string conformance) {
        return SetPdfAIdentification(new PdfAIdentification(part, conformance));
    }

    /// <summary>Sets PDF/UA XMP identification metadata and enables XMP metadata emission.</summary>
    public PdfOptions SetPdfUaIdentification(PdfUaIdentification? identification) {
        PdfUaIdentification = identification;
        if (identification != null) {
            IncludeXmpMetadata = true;
        }

        return this;
    }

    /// <summary>Sets PDF/UA XMP identification metadata and enables XMP metadata emission.</summary>
    public PdfOptions SetPdfUaIdentification(int part = 1) {
        return SetPdfUaIdentification(new PdfUaIdentification(part));
    }

    /// <summary>Sets Factur-X/ZUGFeRD XMP extension metadata and enables XMP metadata emission.</summary>
    public PdfOptions SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata? metadata) {
        ElectronicInvoiceMetadata = metadata;
        if (metadata != null) {
            IncludeXmpMetadata = true;
        }

        return this;
    }

    /// <summary>Sets Factur-X/ZUGFeRD XMP extension metadata and enables XMP metadata emission.</summary>
    public PdfOptions SetElectronicInvoiceMetadata(string conformanceLevel, string version = "1.0") {
        return SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX(conformanceLevel, version));
    }

    /// <summary>Sets a generated catalog output intent backed by an ICC profile.</summary>
    public PdfOptions SetOutputIntent(PdfOutputIntent? outputIntent) {
        OutputIntent = outputIntent;
        return this;
    }

    /// <summary>Sets a generated catalog output intent from ICC profile bytes using the default sRGB output condition identifier.</summary>
    public PdfOptions SetOutputIntent(byte[] iccProfile) {
        return SetOutputIntent(iccProfile, "sRGB IEC61966-2.1", PdfOutputIntentPolicy.Unspecified);
    }

    /// <summary>Sets a generated catalog output intent from ICC profile bytes.</summary>
    public PdfOptions SetOutputIntent(byte[] iccProfile, string outputConditionIdentifier) {
        return SetOutputIntent(iccProfile, outputConditionIdentifier, PdfOutputIntentPolicy.Unspecified);
    }

    /// <summary>Sets a generated catalog output intent from ICC profile bytes using the default sRGB output condition identifier.</summary>
    public PdfOptions SetOutputIntent(byte[] iccProfile, PdfOutputIntentPolicy policy) {
        return SetOutputIntent(iccProfile, "sRGB IEC61966-2.1", policy);
    }

    /// <summary>Sets a generated catalog output intent from ICC profile bytes.</summary>
    public PdfOptions SetOutputIntent(byte[] iccProfile, string outputConditionIdentifier, PdfOutputIntentPolicy policy) {
        OutputIntent = new PdfOutputIntent(iccProfile, outputConditionIdentifier, policy);
        return this;
    }

    /// <summary>Sets the generated catalog output intent to OfficeIMO's built-in sRGB IEC61966-2.1 ICC profile.</summary>
    public PdfOptions SetSrgbOutputIntent() {
        OutputIntent = PdfOutputIntent.CreateSrgbIec6196621();
        return this;
    }

    /// <summary>Sets the generated tagged-PDF groundwork mode.</summary>
    public PdfOptions SetTaggedStructureMode(PdfTaggedStructureMode mode) {
        TaggedStructureMode = mode;
        return this;
    }

    /// <summary>Emits catalog-level tagged-PDF markers without claiming full tagged-content generation.</summary>
    public PdfOptions EnableTaggedPdfCatalogMarkers() {
        TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers;
        return this;
    }

    /// <summary>Sets or clears the generated catalog document language.</summary>
    public PdfOptions SetLanguage(string? language) {
        Language = language;
        return this;
    }

    /// <summary>Sets or clears the generated catalog page mode.</summary>
    public PdfOptions SetCatalogPageMode(PdfCatalogPageMode? pageMode) {
        CatalogPageMode = pageMode;
        return this;
    }

    /// <summary>Sets or clears the generated catalog page layout.</summary>
    public PdfOptions SetCatalogPageLayout(PdfCatalogPageLayout? pageLayout) {
        CatalogPageLayout = pageLayout;
        return this;
    }

    /// <summary>Sets generated catalog page mode and page layout viewer hints.</summary>
    public PdfOptions SetCatalogView(PdfCatalogPageMode? pageMode = null, PdfCatalogPageLayout? pageLayout = null) {
        CatalogPageMode = pageMode;
        CatalogPageLayout = pageLayout;
        return this;
    }

    /// <summary>Clears generated catalog page mode and page layout viewer hints.</summary>
    public PdfOptions ClearCatalogView() {
        _catalogPageMode = null;
        _catalogPageLayout = null;
        return this;
    }

    /// <summary>Sets the generated catalog open action to a page destination.</summary>
    public PdfOptions SetOpenAction(
        int pageNumber = 1,
        double? destinationTop = null,
        PdfOpenActionDestinationMode destinationMode = PdfOpenActionDestinationMode.Xyz,
        double? destinationLeft = null,
        double? destinationBottom = null,
        double? destinationRight = null) {
        OpenAction = new PdfOpenActionOptions(pageNumber, destinationTop, destinationMode, destinationLeft, destinationBottom, destinationRight);
        return this;
    }

    /// <summary>Clears the generated catalog open action.</summary>
    public PdfOptions ClearOpenAction() {
        _openAction = null;
        return this;
    }

    /// <summary>Enables or disables generated catalog page labels.</summary>
    public PdfOptions SetPageLabels(bool include = true, string? prefix = null) {
        IncludePageLabels = include;
        PageLabelPrefix = prefix;
        _pageLabelRanges?.Clear();
        return this;
    }

    /// <summary>Adds a generated catalog page-label rule beginning at the specified one-based document page.</summary>
    public PdfOptions AddPageLabelRange(int startPageNumber, PdfPageNumberStyle style, int startNumber = 1, string? prefix = null) {
        var range = new PdfPageLabelRange(startPageNumber, style, startNumber, prefix);
        if ((_pageLabelRanges ??= new System.Collections.Generic.List<PdfPageLabelRange>()).Any(existing => existing.StartPageNumber == range.StartPageNumber)) {
            throw new ArgumentException("A PDF page-label range already starts at the specified page.", nameof(startPageNumber));
        }

        _pageLabelRanges.Add(range);
        IncludePageLabels = true;
        return this;
    }

    /// <summary>Clears generated catalog page-label ranges while leaving simple page-label options unchanged.</summary>
    public PdfOptions ClearPageLabelRanges() {
        _pageLabelRanges?.Clear();
        return this;
    }

    /// <summary>Enables or disables flattening generated FreeText and Highlight annotations into page content.</summary>
    public PdfOptions SetFlattenVisualAnnotations(bool flatten = true) {
        FlattenVisualAnnotations = flatten;
        return this;
    }

    /// <summary>Sets or clears generated catalog viewer preferences.</summary>
    public PdfOptions SetViewerPreferences(PdfViewerPreferencesOptions? preferences) {
        ViewerPreferences = preferences;
        return this;
    }

    /// <summary>Configures generated catalog viewer preferences.</summary>
    public PdfOptions ConfigureViewerPreferences(System.Action<PdfViewerPreferencesOptions> configure) {
        Guard.NotNull(configure, nameof(configure));
        var preferences = _viewerPreferences?.Clone() ?? new PdfViewerPreferencesOptions();
        configure(preferences);
        _viewerPreferences = preferences;
        return this;
    }

    /// <summary>Sets or clears the generated catalog URI base used by viewers to resolve relative URI actions.</summary>
    public PdfOptions SetCatalogUriBase(string? uriBase) {
        CatalogUriBase = uriBase;
        return this;
    }

    /// <summary>Clears the generated catalog URI base.</summary>
    public PdfOptions ClearCatalogUriBase() {
        _catalogUriBase = null;
        return this;
    }

    /// <summary>Sets or clears the generated AcroForm default text alignment emitted through /Q.</summary>
    public PdfOptions SetAcroFormDefaultTextAlignment(PdfFormFieldTextAlignment? alignment) {
        AcroFormDefaultTextAlignment = alignment;
        return this;
    }

    /// <summary>Clears the generated AcroForm default text alignment.</summary>
    public PdfOptions ClearAcroFormDefaultTextAlignment() {
        _acroFormDefaultTextAlignment = null;
        return this;
    }

    /// <summary>
    /// Configures common PDF/UA-1 groundwork without enabling formal compliance profile generation.
    /// </summary>
    public PdfOptions ConfigurePdfUaGroundwork(string language = "en-US") {
        return ConfigurePdfUaGroundwork(PdfComplianceProfile.PdfUa1, language);
    }

    /// <summary>
    /// Applies common PDF/UA groundwork for the requested profile.
    /// </summary>
    public PdfOptions UsePdfUa(PdfComplianceProfile profile = PdfComplianceProfile.PdfUa1, string language = "en-US") {
        return ConfigurePdfUaGroundwork(profile, language);
    }

    /// <summary>
    /// Configures common PDF/UA-1 or PDF/UA-2 groundwork without enabling formal compliance profile generation.
    /// </summary>
    public PdfOptions ConfigurePdfUaGroundwork(PdfComplianceProfile profile, string language = "en-US") {
        if (profile != PdfComplianceProfile.PdfUa1 && profile != PdfComplianceProfile.PdfUa2) {
            throw new System.ArgumentException("PDF/UA groundwork requires a PDF/UA-1 or PDF/UA-2 compliance profile.", nameof(profile));
        }

        Language = language;
        FileVersion = profile == PdfComplianceProfile.PdfUa2 ? PdfFileVersion.Pdf20 : PdfFileVersion.Pdf17;
        IncludeStandardFontToUnicodeMaps = true;
        SetPdfUaIdentification(profile == PdfComplianceProfile.PdfUa2 ? 2 : 1);
        EnableTaggedPdfCatalogMarkers();

        var preferences = _viewerPreferences?.Clone() ?? new PdfViewerPreferencesOptions();
        preferences.DisplayDocTitle = true;
        _viewerPreferences = preferences;
        return this;
    }

    /// <summary>
    /// Configures common PDF/A-2, PDF/A-3, or PDF/A-4 groundwork without enabling formal compliance profile generation.
    /// </summary>
    public PdfOptions ConfigurePdfAGroundwork(PdfComplianceProfile profile, string language = "en-US") {
        PdfAIdentification identification = CreatePdfAIdentification(profile);
        FileVersion = profile == PdfComplianceProfile.PdfA4 || profile == PdfComplianceProfile.PdfA4E || profile == PdfComplianceProfile.PdfA4F
            ? PdfFileVersion.Pdf20
            : PdfFileVersion.Pdf17;
        SetPdfAIdentification(identification);
        SetSrgbOutputIntent();

        if (RequiresPdfAUnicodeGroundwork(profile)) {
            IncludeStandardFontToUnicodeMaps = true;
        }

        if (RequiresPdfAAccessibilityGroundwork(profile)) {
            Language = language;
            EnableTaggedPdfCatalogMarkers();
        }

        return this;
    }

    /// <summary>
    /// Applies common PDF/A groundwork for the requested profile.
    /// </summary>
    public PdfOptions UsePdfA(PdfComplianceProfile profile = PdfComplianceProfile.PdfA3B, string language = "en-US") {
        return ConfigurePdfAGroundwork(profile, language);
    }

    /// <summary>Adds an embedded file associated with the generated PDF catalog.</summary>
    public PdfOptions AddEmbeddedFile(PdfEmbeddedFile file) {
        Guard.NotNull(file, nameof(file));
        if (_embeddedFiles != null) {
            foreach (PdfEmbeddedFile existingFile in _embeddedFiles) {
                if (string.Equals(existingFile.FileName, file.FileName, System.StringComparison.Ordinal)) {
                    throw new System.ArgumentException("PDF embedded file names must be unique within a generated document.", nameof(file));
                }
            }
        }

        (_embeddedFiles ??= new System.Collections.Generic.List<PdfEmbeddedFile>()).Add(file.Clone());
        return this;
    }

    /// <summary>Adds an embedded file associated with the generated PDF catalog.</summary>
    public PdfOptions AddEmbeddedFile(
        string fileName,
        byte[] data,
        string? mimeType = null,
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Unspecified,
        string? description = null) {
        return AddEmbeddedFile(new PdfEmbeddedFile(fileName, data, mimeType, relationship, description));
    }

    /// <summary>
    /// Adds the canonical Factur-X/ZUGFeRD CrossIndustryInvoice XML payload and matching XMP extension metadata.
    /// </summary>
    public PdfOptions AddFacturXInvoiceXml(
        byte[] ciiXml,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML") {
        PdfElectronicInvoiceMetadata metadata = CreateFacturXInvoiceMetadata(conformanceLevel, version);
        PdfEmbeddedFile attachment = CreateFacturXInvoiceAttachment(ciiXml, relationship, description);
        AddEmbeddedFile(attachment);
        return SetElectronicInvoiceMetadata(metadata);
    }

    /// <summary>
    /// Adds the canonical Factur-X/ZUGFeRD CrossIndustryInvoice XML file and matching XMP extension metadata.
    /// </summary>
    public PdfOptions AddFacturXInvoiceXmlFile(
        string ciiXmlPath,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML") {
        Guard.NotNullOrWhiteSpace(ciiXmlPath, nameof(ciiXmlPath));
        return AddFacturXInvoiceXml(System.IO.File.ReadAllBytes(ciiXmlPath), conformanceLevel, version, relationship, description);
    }

    /// <summary>
    /// Configures common PDF/A-3 Factur-X/ZUGFeRD groundwork without enabling formal compliance profile generation.
    /// </summary>
    public PdfOptions ConfigureFacturXGroundwork(
        byte[] ciiXml,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML",
        bool useDocumentFontFallback = true) {
        return ConfigureFacturXGroundwork(
            ciiXml,
            useDocumentFontFallback ? PdfTextFallbackFeatures.DocumentFont : PdfTextFallbackFeatures.None,
            conformanceLevel,
            version,
            relationship,
            description);
    }

    /// <summary>
    /// Configures common PDF/A-3 Factur-X/ZUGFeRD groundwork with explicit text fallback groups.
    /// </summary>
    public PdfOptions ConfigureFacturXGroundwork(
        byte[] ciiXml,
        PdfTextFallbackFeatures textFallbacks,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML") {
        PdfAIdentification pdfAIdentification = new PdfAIdentification(3, "B");
        PdfOutputIntent outputIntent = PdfOutputIntent.CreateSrgbIec6196621();
        PdfElectronicInvoiceMetadata metadata = CreateFacturXInvoiceMetadata(conformanceLevel, version);
        PdfEmbeddedFile attachment = CreateFacturXInvoiceAttachment(ciiXml, relationship, description);

        AddEmbeddedFile(attachment);
        FileVersion = PdfFileVersion.Pdf17;
        IncludeStandardFontToUnicodeMaps = true;
        if (textFallbacks != PdfTextFallbackFeatures.None) {
            UseTextFallbacks(textFallbacks);
        }

        SetPdfAIdentification(pdfAIdentification);
        SetOutputIntent(outputIntent);
        return SetElectronicInvoiceMetadata(metadata);
    }

    /// <summary>
    /// Applies Factur-X/ZUGFeRD PDF/A-3 groundwork and attaches the CrossIndustryInvoice XML payload.
    /// </summary>
    public PdfOptions UseFacturX(
        byte[] ciiXml,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML",
        PdfTextFallbackFeatures textFallbacks = PdfTextFallbackFeatures.DocumentFont) {
        return ConfigureFacturXGroundwork(ciiXml, textFallbacks, conformanceLevel, version, relationship, description);
    }

    /// <summary>
    /// Configures common PDF/A-3 Factur-X/ZUGFeRD groundwork from a CrossIndustryInvoice XML file.
    /// </summary>
    public PdfOptions ConfigureFacturXGroundworkFile(
        string ciiXmlPath,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML",
        bool useDocumentFontFallback = true) {
        Guard.NotNullOrWhiteSpace(ciiXmlPath, nameof(ciiXmlPath));
        return ConfigureFacturXGroundwork(System.IO.File.ReadAllBytes(ciiXmlPath), conformanceLevel, version, relationship, description, useDocumentFontFallback);
    }

    /// <summary>
    /// Applies Factur-X/ZUGFeRD PDF/A-3 groundwork from a CrossIndustryInvoice XML file.
    /// </summary>
    public PdfOptions UseFacturXFile(
        string ciiXmlPath,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML",
        PdfTextFallbackFeatures textFallbacks = PdfTextFallbackFeatures.DocumentFont) {
        Guard.NotNullOrWhiteSpace(ciiXmlPath, nameof(ciiXmlPath));
        return UseFacturX(System.IO.File.ReadAllBytes(ciiXmlPath), conformanceLevel, version, relationship, description, textFallbacks);
    }

    /// <summary>
    /// Configures common PDF/A-3 e-invoice groundwork for Factur-X or ZUGFeRD without enabling formal compliance profile generation.
    /// </summary>
    public PdfOptions ConfigureElectronicInvoiceGroundwork(
        PdfComplianceProfile profile,
        byte[] ciiXml,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML",
        bool useDocumentFontFallback = true) {
        ValidateElectronicInvoiceProfile(profile);
        return ConfigureFacturXGroundwork(ciiXml, conformanceLevel, version, relationship, description, useDocumentFontFallback);
    }

    /// <summary>
    /// Configures common PDF/A-3 e-invoice groundwork for Factur-X or ZUGFeRD from a CrossIndustryInvoice XML file.
    /// </summary>
    public PdfOptions ConfigureElectronicInvoiceGroundworkFile(
        PdfComplianceProfile profile,
        string ciiXmlPath,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML",
        bool useDocumentFontFallback = true) {
        Guard.NotNullOrWhiteSpace(ciiXmlPath, nameof(ciiXmlPath));
        return ConfigureElectronicInvoiceGroundwork(profile, System.IO.File.ReadAllBytes(ciiXmlPath), conformanceLevel, version, relationship, description, useDocumentFontFallback);
    }

    /// <summary>Removes all embedded files associated with the generated PDF catalog.</summary>
    public PdfOptions ClearEmbeddedFiles() {
        _embeddedFiles?.Clear();
        return this;
    }

    internal bool TryGetEmbeddedStandardFont(PdfStandardFont font, out PdfEmbeddedFont? embeddedFont) {
        if (_embeddedFonts != null && _embeddedFonts.TryGetValue(font, out PdfEmbeddedFont? value)) {
            embeddedFont = value.Clone();
            return true;
        }

        embeddedFont = null;
        return false;
    }

    internal bool TryGetEmbeddedStandardFontProgram(PdfStandardFont font, out PdfTrueTypeFontProgram? fontProgram) {
        Guard.StandardFont(font, nameof(font), "PDF embedded font lookup must target one of the supported standard PDF fonts.");
        if (_embeddedFontPrograms != null && _embeddedFontPrograms.TryGetValue(font, out PdfTrueTypeFontProgram? cachedProgram)) {
            fontProgram = cachedProgram;
            return true;
        }

        if (_embeddedFontProgramFailures != null && _embeddedFontProgramFailures.Contains(font)) {
            fontProgram = null;
            return false;
        }

        if (_embeddedFonts == null || !_embeddedFonts.TryGetValue(font, out PdfEmbeddedFont? embeddedFont)) {
            fontProgram = null;
            return false;
        }

        if (IsOpenTypeCffFontData(embeddedFont.DataSnapshot)) {
            fontProgram = null;
            return false;
        }

        try {
            fontProgram = PdfTrueTypeFontProgram.Parse(embeddedFont.DataSnapshot, embeddedFont.FontName);
        } catch (System.Exception exception) when (PdfFontDiagnostics.IsFontProgramException(exception)) {
            ReportEmbeddedFontProgramFailure(font, embeddedFont, exception);
            (_embeddedFontProgramFailures ??= new System.Collections.Generic.HashSet<PdfStandardFont>()).Add(font);
            fontProgram = null;
            return false;
        }

        (_embeddedFontPrograms ??= new System.Collections.Generic.Dictionary<PdfStandardFont, PdfTrueTypeFontProgram>())[font] = fontProgram;
        return true;
    }

    internal bool TryGetEmbeddedStandardFontProgramForGeneration(PdfStandardFont font, out PdfEmbeddedFont? embeddedFont, out PdfTrueTypeFontProgram? fontProgram) {
        Guard.StandardFont(font, nameof(font), "PDF embedded font lookup must target one of the supported standard PDF fonts.");
        embeddedFont = null;
        if (_embeddedFonts == null || !_embeddedFonts.TryGetValue(font, out embeddedFont)) {
            fontProgram = null;
            return false;
        }

        if (IsOpenTypeCffFontData(embeddedFont.DataSnapshot)) {
            fontProgram = null;
            return false;
        }

        if (_embeddedFontPrograms != null && _embeddedFontPrograms.TryGetValue(font, out PdfTrueTypeFontProgram? cachedProgram)) {
            fontProgram = cachedProgram;
            return true;
        }

        try {
            fontProgram = PdfTrueTypeFontProgram.Parse(embeddedFont.DataSnapshot, embeddedFont.FontName);
        } catch (System.Exception exception) when (PdfFontDiagnostics.IsFontProgramException(exception)) {
            ReportEmbeddedFontProgramFailure(font, embeddedFont, exception);
            throw;
        }

        (_embeddedFontPrograms ??= new System.Collections.Generic.Dictionary<PdfStandardFont, PdfTrueTypeFontProgram>())[font] = fontProgram;
        _embeddedFontProgramFailures?.Remove(font);
        ClearReportedEmbeddedFontProgramFailure(font);
        return true;
    }

    internal bool TryGetEmbeddedStandardOpenTypeCffFontProgram(PdfStandardFont font, out PdfOpenTypeCffFontProgram? fontProgram) {
        Guard.StandardFont(font, nameof(font), "PDF embedded font lookup must target one of the supported standard PDF fonts.");
        if (_embeddedOpenTypeCffFontPrograms != null && _embeddedOpenTypeCffFontPrograms.TryGetValue(font, out PdfOpenTypeCffFontProgram? cachedProgram)) {
            fontProgram = cachedProgram;
            return true;
        }

        if (_embeddedFontProgramFailures != null && _embeddedFontProgramFailures.Contains(font)) {
            fontProgram = null;
            return false;
        }

        if (_embeddedFonts == null || !_embeddedFonts.TryGetValue(font, out PdfEmbeddedFont? embeddedFont) || !IsOpenTypeCffFontData(embeddedFont.DataSnapshot)) {
            fontProgram = null;
            return false;
        }

        try {
            fontProgram = PdfOpenTypeCffFontProgram.Parse(embeddedFont.DataSnapshot, embeddedFont.FontName);
        } catch (System.Exception exception) when (PdfFontDiagnostics.IsFontProgramException(exception)) {
            ReportEmbeddedFontProgramFailure(font, embeddedFont, exception);
            (_embeddedFontProgramFailures ??= new System.Collections.Generic.HashSet<PdfStandardFont>()).Add(font);
            fontProgram = null;
            return false;
        }

        (_embeddedOpenTypeCffFontPrograms ??= new System.Collections.Generic.Dictionary<PdfStandardFont, PdfOpenTypeCffFontProgram>())[font] = fontProgram;
        return true;
    }

    internal bool TryGetEmbeddedStandardOpenTypeCffFontProgramForGeneration(PdfStandardFont font, out PdfEmbeddedFont? embeddedFont, out PdfOpenTypeCffFontProgram? fontProgram) {
        Guard.StandardFont(font, nameof(font), "PDF embedded font lookup must target one of the supported standard PDF fonts.");
        embeddedFont = null;
        if (_embeddedFonts == null || !_embeddedFonts.TryGetValue(font, out embeddedFont) || !IsOpenTypeCffFontData(embeddedFont.DataSnapshot)) {
            fontProgram = null;
            return false;
        }

        if (_embeddedOpenTypeCffFontPrograms != null && _embeddedOpenTypeCffFontPrograms.TryGetValue(font, out PdfOpenTypeCffFontProgram? cachedProgram)) {
            fontProgram = cachedProgram;
            return true;
        }

        try {
            fontProgram = PdfOpenTypeCffFontProgram.Parse(embeddedFont.DataSnapshot, embeddedFont.FontName);
        } catch (System.Exception exception) when (PdfFontDiagnostics.IsFontProgramException(exception)) {
            ReportEmbeddedFontProgramFailure(font, embeddedFont, exception);
            throw;
        }

        (_embeddedOpenTypeCffFontPrograms ??= new System.Collections.Generic.Dictionary<PdfStandardFont, PdfOpenTypeCffFontProgram>())[font] = fontProgram;
        _embeddedFontProgramFailures?.Remove(font);
        ClearReportedEmbeddedFontProgramFailure(font);
        return true;
    }

    private void ReportEmbeddedFontProgramFailure(PdfStandardFont font, PdfEmbeddedFont embeddedFont, System.Exception exception) {
        string source = "embedded-font:" + font;
        AddFontDiagnostics(font, PdfFontDiagnostics.AnalyzeEmbeddedFontFailure(embeddedFont.DataSnapshot, source, embeddedFont.FontName, exception));
    }

    private void ClearReportedEmbeddedFontProgramFailure(PdfStandardFont font) {
        if (_reportedEmbeddedFontProgramFailures == null || _reportedEmbeddedFontProgramFailures.Count == 0) {
            return;
        }

        string prefix = font.ToString() + "|";
        _reportedEmbeddedFontProgramFailures.RemoveWhere(key => key.StartsWith(prefix, System.StringComparison.Ordinal));
    }

    internal void ResetEmbeddedFontProgramUsage() {
        if (_embeddedFontPrograms != null) {
            foreach (PdfTrueTypeFontProgram program in _embeddedFontPrograms.Values) {
                program.ResetGlyphUsage();
            }
        }

        if (_embeddedOpenTypeCffFontPrograms != null) {
            foreach (PdfOpenTypeCffFontProgram program in _embeddedOpenTypeCffFontPrograms.Values) {
                program.ResetGlyphUsage();
            }
        }

        ResetNamedFontProgramUsage();
    }

    private static bool IsOpenTypeCffFontData(byte[] fontData) =>
        fontData.Length >= 4 &&
        fontData[0] == 0x4F &&
        fontData[1] == 0x54 &&
        fontData[2] == 0x54 &&
        fontData[3] == 0x4F;

    private static PdfElectronicInvoiceMetadata CreateFacturXInvoiceMetadata(string conformanceLevel, string version) {
        return PdfElectronicInvoiceMetadata.FacturX(conformanceLevel, version);
    }

    private static PdfAIdentification CreatePdfAIdentification(PdfComplianceProfile profile) {
        switch (profile) {
            case PdfComplianceProfile.PdfA2B:
                return new PdfAIdentification(2, "B");
            case PdfComplianceProfile.PdfA2U:
                return new PdfAIdentification(2, "U");
            case PdfComplianceProfile.PdfA2A:
                return new PdfAIdentification(2, "A");
            case PdfComplianceProfile.PdfA3B:
                return new PdfAIdentification(3, "B");
            case PdfComplianceProfile.PdfA3U:
                return new PdfAIdentification(3, "U");
            case PdfComplianceProfile.PdfA3A:
                return new PdfAIdentification(3, "A");
            case PdfComplianceProfile.PdfA4:
                return PdfAIdentification.PdfA4();
            case PdfComplianceProfile.PdfA4E:
                return PdfAIdentification.PdfA4E();
            case PdfComplianceProfile.PdfA4F:
                return PdfAIdentification.PdfA4F();
            default:
                throw new System.ArgumentException("PDF/A groundwork requires a PDF/A-2, PDF/A-3, or PDF/A-4 compliance profile.", nameof(profile));
        }
    }

    private static bool RequiresPdfAUnicodeGroundwork(PdfComplianceProfile profile) {
        return profile == PdfComplianceProfile.PdfA2U ||
            profile == PdfComplianceProfile.PdfA2A ||
            profile == PdfComplianceProfile.PdfA3U ||
            profile == PdfComplianceProfile.PdfA3A ||
            profile == PdfComplianceProfile.PdfA4 ||
            profile == PdfComplianceProfile.PdfA4E ||
            profile == PdfComplianceProfile.PdfA4F;
    }

    private static bool RequiresPdfAAccessibilityGroundwork(PdfComplianceProfile profile) {
        return profile == PdfComplianceProfile.PdfA2A ||
            profile == PdfComplianceProfile.PdfA3A;
    }

    private static PdfEmbeddedFile CreateFacturXInvoiceAttachment(
        byte[] ciiXml,
        PdfAssociatedFileRelationship relationship,
        string? description) {
        Guard.NotNullOrEmpty(ciiXml, nameof(ciiXml));
        ValidateFacturXInvoiceRelationship(relationship);
        return new PdfEmbeddedFile("factur-x.xml", ciiXml, "application/xml", relationship, description);
    }

    private static void ValidateFacturXInvoiceRelationship(PdfAssociatedFileRelationship relationship) {
        if (relationship != PdfAssociatedFileRelationship.Alternative &&
            relationship != PdfAssociatedFileRelationship.Data &&
            relationship != PdfAssociatedFileRelationship.Source) {
            throw new System.ArgumentException("Factur-X/ZUGFeRD invoice XML must use Alternative, Data, or Source associated-file relationship.", nameof(relationship));
        }
    }

    private static void ValidateElectronicInvoiceProfile(PdfComplianceProfile profile) {
        if (profile != PdfComplianceProfile.FacturX &&
            profile != PdfComplianceProfile.Zugferd) {
            throw new System.ArgumentException("PDF e-invoice groundwork requires FacturX or Zugferd compliance profile.", nameof(profile));
        }
    }
}

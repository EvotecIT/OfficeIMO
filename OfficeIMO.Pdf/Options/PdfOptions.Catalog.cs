namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>When true, generated page content streams are written with Flate compression.</summary>
    public bool CompressContentStreams { get; set; }
    /// <summary>When true, generated standard-font resources include WinAnsi-to-Unicode CMaps for stronger extraction and compliance groundwork.</summary>
    public bool IncludeStandardFontToUnicodeMaps { get; set; }
    /// <summary>When true, embedded TrueType font file streams are Flate-compressed while preserving their original /Length1 metadata.</summary>
    public bool CompressEmbeddedFonts { get; set; } = true;
    /// <summary>When true, generated PDFs include a catalog XMP metadata stream synchronized with document Info metadata.</summary>
    public bool IncludeXmpMetadata { get; set; }
    /// <summary>When true, generated PDFs include catalog page labels that match the configured page-number style and start number.</summary>
    public bool IncludePageLabels { get; set; }
    /// <summary>Optional catalog page-label prefix, for example "A-" or "Appendix ". Requires <see cref="IncludePageLabels"/> to be emitted.</summary>
    public string? PageLabelPrefix {
        get => _pageLabelPrefix;
        set {
            PdfPageLabelDictionaryBuilder.ValidatePrefix(value, nameof(PageLabelPrefix));
            _pageLabelPrefix = value;
        }
    }
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

    /// <summary>Optional simple viewer preferences emitted through the generated catalog.</summary>
    public PdfViewerPreferencesOptions? ViewerPreferences {
        get => _viewerPreferences?.Clone();
        set => _viewerPreferences = value?.Clone();
    }

    internal PdfViewerPreferencesOptions? ViewerPreferencesSnapshot => _viewerPreferences?.Clone();

    /// <summary>Embedded files associated with the generated document catalog.</summary>
    public System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile> EmbeddedFiles =>
        (System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile>?)CloneEmbeddedFiles(_embeddedFiles) ?? System.Array.Empty<PdfEmbeddedFile>();

    internal System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile> EmbeddedFileSnapshots =>
        (System.Collections.Generic.IReadOnlyList<PdfEmbeddedFile>?)CloneEmbeddedFiles(_embeddedFiles) ?? System.Array.Empty<PdfEmbeddedFile>();

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
        _embeddedFontProgramFailures?.Remove(font);
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
        FooterFont = PdfStandardFont.Helvetica;
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
        _embeddedFontProgramFailures?.Clear();
        return this;
    }

    internal PdfOptions UseDefaultTextFontFamily(PdfEmbeddedFontFamily fontFamily) {
        Guard.NotNull(fontFamily, nameof(fontFamily));
        PdfEmbeddedFontFamily snapshot = fontFamily.Clone();

        EmbedStandardFont(PdfStandardFont.Helvetica, snapshot.RegularSnapshot, BuildFontFamilyFaceName(snapshot.FamilyName, "Regular"));
        EmbedStandardFont(PdfStandardFont.HelveticaBold, snapshot.BoldSnapshot ?? snapshot.RegularSnapshot, BuildFontFamilyFaceName(snapshot.FamilyName, "Bold"));
        EmbedStandardFont(PdfStandardFont.HelveticaOblique, snapshot.ItalicSnapshot ?? snapshot.RegularSnapshot, BuildFontFamilyFaceName(snapshot.FamilyName, "Italic"));
        EmbedStandardFont(PdfStandardFont.HelveticaBoldOblique, snapshot.BoldItalicSnapshot ?? snapshot.BoldSnapshot ?? snapshot.ItalicSnapshot ?? snapshot.RegularSnapshot, BuildFontFamilyFaceName(snapshot.FamilyName, "BoldItalic"));
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

    /// <summary>Enables or disables generated catalog page labels.</summary>
    public PdfOptions SetPageLabels(bool include = true, string? prefix = null) {
        IncludePageLabels = include;
        PageLabelPrefix = prefix;
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

    /// <summary>
    /// Configures common PDF/UA-1 groundwork without enabling formal compliance profile generation.
    /// </summary>
    public PdfOptions ConfigurePdfUaGroundwork(string language = "en-US") {
        Language = language;
        FileVersion = PdfFileVersion.Pdf17;
        IncludeStandardFontToUnicodeMaps = true;
        SetPdfUaIdentification();
        EnableTaggedPdfCatalogMarkers();

        var preferences = _viewerPreferences?.Clone() ?? new PdfViewerPreferencesOptions();
        preferences.DisplayDocTitle = true;
        _viewerPreferences = preferences;
        return this;
    }

    /// <summary>
    /// Configures common PDF/A-2 or PDF/A-3 groundwork without enabling formal compliance profile generation.
    /// </summary>
    public PdfOptions ConfigurePdfAGroundwork(PdfComplianceProfile profile, string language = "en-US") {
        PdfAIdentification identification = CreatePdfAIdentification(profile);
        FileVersion = PdfFileVersion.Pdf17;
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
    /// Configures common PDF/A-3 Factur-X/ZUGFeRD groundwork without enabling formal compliance profile generation.
    /// </summary>
    public PdfOptions ConfigureFacturXGroundwork(
        byte[] ciiXml,
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
        SetPdfAIdentification(pdfAIdentification);
        SetOutputIntent(outputIntent);
        return SetElectronicInvoiceMetadata(metadata);
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

        try {
            fontProgram = PdfTrueTypeFontProgram.Parse(embeddedFont.DataSnapshot, embeddedFont.FontName);
        } catch (System.NotSupportedException) {
            (_embeddedFontProgramFailures ??= new System.Collections.Generic.HashSet<PdfStandardFont>()).Add(font);
            fontProgram = null;
            return false;
        }

        (_embeddedFontPrograms ??= new System.Collections.Generic.Dictionary<PdfStandardFont, PdfTrueTypeFontProgram>())[font] = fontProgram;
        return true;
    }

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
            default:
                throw new System.ArgumentException("PDF/A groundwork requires a PDF/A-2 or PDF/A-3 compliance profile.", nameof(profile));
        }
    }

    private static bool RequiresPdfAUnicodeGroundwork(PdfComplianceProfile profile) {
        return profile == PdfComplianceProfile.PdfA2U ||
            profile == PdfComplianceProfile.PdfA2A ||
            profile == PdfComplianceProfile.PdfA3U ||
            profile == PdfComplianceProfile.PdfA3A;
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
}

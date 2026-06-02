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
    /// <summary>Requested generated-PDF compliance profile. Non-None profiles are validated strictly and fail until their required primitives are implemented.</summary>
    public PdfComplianceProfile ComplianceProfile {
        get => _complianceProfile;
        set {
            Guard.ComplianceProfile(value, nameof(ComplianceProfile));
            _complianceProfile = value;
        }
    }
    /// <summary>Optional generated catalog output intent backed by an ICC profile.</summary>
    public PdfOutputIntent? OutputIntent {
        get => _outputIntent?.Clone();
        set => _outputIntent = value?.Clone();
    }
    internal PdfOutputIntent? OutputIntentSnapshot => _outputIntent?.Clone();
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
        return this;
    }

    /// <summary>Embeds a TrueType font file from disk for a generated standard-font slot.</summary>
    public PdfOptions EmbedStandardFont(PdfStandardFont font, string path, string? fontName = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return EmbedStandardFont(font, System.IO.File.ReadAllBytes(path), fontName);
    }

    /// <summary>Removes all embedded standard-font mappings.</summary>
    public PdfOptions ClearEmbeddedStandardFonts() {
        _embeddedFonts?.Clear();
        return this;
    }

    /// <summary>Requests a generated-PDF compliance profile for this document.</summary>
    public PdfOptions RequireCompliance(PdfComplianceProfile profile) {
        ComplianceProfile = profile;
        return this;
    }

    /// <summary>Sets a generated catalog output intent backed by an ICC profile.</summary>
    public PdfOptions SetOutputIntent(PdfOutputIntent? outputIntent) {
        OutputIntent = outputIntent;
        return this;
    }

    /// <summary>Sets a generated catalog output intent from ICC profile bytes.</summary>
    public PdfOptions SetOutputIntent(byte[] iccProfile, string outputConditionIdentifier = "sRGB IEC61966-2.1") {
        OutputIntent = new PdfOutputIntent(iccProfile, outputConditionIdentifier);
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
}

using OfficeIMO.Html;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Options for exporting PowerPoint presentations to HTML.
/// </summary>
public sealed class PowerPointHtmlSaveOptions {
    private PowerPointHtmlExportProfile _exportProfile = PowerPointHtmlExportProfile.SemanticSlides;
    private OfficeHtmlDocumentOptions _documentOutput = new() {
        Title = null,
        Theme = OfficeVisualThemeKind.WordLike,
        BodyClass = "officeimo-html officeimo-powerpoint-html"
    };

    /// <summary>PowerPoint-to-HTML lane to export. Defaults to semantic slides.</summary>
    public PowerPointHtmlExportProfile ExportProfile {
        get => _exportProfile;
        set {
            if (!Enum.IsDefined(typeof(PowerPointHtmlExportProfile), value)) {
                throw new ArgumentOutOfRangeException(nameof(value), value, "PowerPoint HTML export profile is not supported.");
            }
            _exportProfile = value;
        }
    }

    /// <summary>
    /// Compatibility bridge to the former cross-format profile enum. New code should use
    /// <see cref="ExportProfile"/> so only PowerPoint profiles are representable.
    /// </summary>
    public OfficeHtmlConversionProfile Profile {
        get => ExportProfile == PowerPointHtmlExportProfile.VisualReview
            ? OfficeHtmlConversionProfile.PowerPointVisualReview
            : OfficeHtmlConversionProfile.PowerPointSemanticSlides;
        set => ExportProfile = value switch {
            OfficeHtmlConversionProfile.PowerPointSemanticSlides => PowerPointHtmlExportProfile.SemanticSlides,
            OfficeHtmlConversionProfile.PowerPointVisualReview => PowerPointHtmlExportProfile.VisualReview,
            _ => throw new ArgumentOutOfRangeException(nameof(value), value, "The selected HTML conversion profile is not a PowerPoint profile.")
        };
    }

    /// <summary>Shared engine profile used by the selected PowerPoint export lane.</summary>
    public HtmlConversionProfile SharedProfile => ExportProfile == PowerPointHtmlExportProfile.VisualReview
        ? HtmlConversionProfile.PositionedReview
        : HtmlConversionProfile.Semantic;

    /// <summary>Creates semantic slide export settings.</summary>
    public static PowerPointHtmlSaveOptions CreateSemanticSlidesProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.WordLike) => new() {
        ExportProfile = PowerPointHtmlExportProfile.SemanticSlides,
        Theme = theme
    };

    /// <summary>Creates positioned visual-review export settings.</summary>
    public static PowerPointHtmlSaveOptions CreateVisualReviewProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.Report) => new() {
        ExportProfile = PowerPointHtmlExportProfile.VisualReview,
        Theme = theme
    };

    /// <summary>Composed document-versus-fragment, theme, title, language, style, and newline settings.</summary>
    public OfficeHtmlDocumentOptions DocumentOutput {
        get => _documentOutput;
        set => _documentOutput = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Compatibility alias for <see cref="OfficeHtmlDocumentOptions.Theme"/>.</summary>
    public OfficeVisualThemeKind Theme { get => DocumentOutput.Theme; set => DocumentOutput.Theme = value; }

    /// <summary>Compatibility alias for <see cref="OfficeHtmlDocumentOptions.IncludeDefaultStyles"/>.</summary>
    public bool IncludeDefaultStyles { get => DocumentOutput.IncludeDefaultStyles; set => DocumentOutput.IncludeDefaultStyles = value; }

    /// <summary>Compatibility alias for <see cref="OfficeHtmlDocumentOptions.Title"/>.</summary>
    public string? Title { get => DocumentOutput.Title; set => DocumentOutput.Title = value; }

    /// <summary>Compatibility alias for <see cref="OfficeHtmlDocumentOptions.EmitDocumentShell"/>.</summary>
    public bool EmitDocumentShell { get => DocumentOutput.EmitDocumentShell; set => DocumentOutput.EmitDocumentShell = value; }

    /// <summary>Compatibility alias for <see cref="OfficeHtmlDocumentOptions.Language"/>.</summary>
    public string? Language { get => DocumentOutput.Language; set => DocumentOutput.Language = value; }

    /// <summary>Compatibility alias for <see cref="OfficeHtmlDocumentOptions.NewLine"/>.</summary>
    public string NewLine { get => DocumentOutput.NewLine; set => DocumentOutput.NewLine = value; }

    /// <summary>When true, hidden slides are included.</summary>
    public bool IncludeHiddenSlides { get; set; }

    /// <summary>When true, notes are included in the extraction proof block.</summary>
    public bool IncludeNotes { get; set; } = true;

    /// <summary>When true, tables are exported.</summary>
    public bool IncludeTables { get; set; } = true;

    /// <summary>When true, hidden shapes are included in semantic and positioned review output.</summary>
    public bool IncludeHiddenShapes { get; set; }

    /// <summary>When true, emits slide-aligned extraction markdown as proof text.</summary>
    public bool IncludeExtractionProof { get; set; } = true;

    /// <summary>When true, exposes masters, layouts, theme identity, and template assets as inert review metadata.</summary>
    public bool IncludeMasterInventory { get; set; } = true;

    /// <summary>When true, emits semantic SmartArt snapshots or diagnosed text fallbacks.</summary>
    public bool IncludeSmartArt { get; set; } = true;

    /// <summary>When true, emits media kind, poster-frame, and playback metadata without executing media.</summary>
    public bool IncludeMedia { get; set; } = true;

    /// <summary>When true, reports supported picture adjustments and diagnosed effect simplifications.</summary>
    public bool IncludeAdvancedEffects { get; set; } = true;

    /// <summary>Creates a reusable copy of these conversion settings.</summary>
    public PowerPointHtmlSaveOptions Clone() => new PowerPointHtmlSaveOptions {
        ExportProfile = ExportProfile,
        DocumentOutput = DocumentOutput.Clone(),
        IncludeHiddenSlides = IncludeHiddenSlides,
        IncludeNotes = IncludeNotes,
        IncludeTables = IncludeTables,
        IncludeHiddenShapes = IncludeHiddenShapes,
        IncludeExtractionProof = IncludeExtractionProof,
        IncludeMasterInventory = IncludeMasterInventory,
        IncludeSmartArt = IncludeSmartArt,
        IncludeMedia = IncludeMedia,
        IncludeAdvancedEffects = IncludeAdvancedEffects
    };

    internal void Validate() {
        if (!Enum.IsDefined(typeof(PowerPointHtmlExportProfile), ExportProfile)) {
            throw new ArgumentOutOfRangeException(nameof(ExportProfile), ExportProfile, "PowerPoint HTML export profile is not supported.");
        }
        DocumentOutput.Validate();
    }
}

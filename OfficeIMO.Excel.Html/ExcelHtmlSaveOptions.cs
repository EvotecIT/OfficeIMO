using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

/// <summary>
/// Options for exporting Excel workbooks and worksheets to HTML.
/// </summary>
public sealed class ExcelHtmlSaveOptions {
    private ExcelHtmlExportProfile _exportProfile = ExcelHtmlExportProfile.SemanticTables;
    private OfficeHtmlDocumentOptions _documentOutput = new() {
        Title = null,
        Theme = OfficeVisualThemeKind.WordLike,
        BodyClass = "officeimo-html officeimo-excel-html"
    };

    /// <summary>Default maximum worksheet rows projected to semantic HTML.</summary>
    public const int DefaultMaxRowsPerSheet = 10000;

    /// <summary>Default maximum worksheet columns projected to semantic HTML.</summary>
    public const int DefaultMaxColumnsPerSheet = 1024;

    /// <summary>Default maximum worksheet cells visited while projecting semantic HTML.</summary>
    public const int DefaultMaxCellsPerSheet = 1000000;

    /// <summary>Default maximum merged-range records inspected per worksheet.</summary>
    public const int DefaultMaxMergedRangesPerSheet = 10000;

    /// <summary>Excel-to-HTML lane to export. Defaults to semantic worksheet tables.</summary>
    public ExcelHtmlExportProfile ExportProfile {
        get => _exportProfile;
        set {
            if (!Enum.IsDefined(typeof(ExcelHtmlExportProfile), value)) {
                throw new ArgumentOutOfRangeException(nameof(value), value, "Excel HTML export profile is not supported.");
            }
            _exportProfile = value;
        }
    }

    /// <summary>
    /// Compatibility bridge to the former cross-format profile enum. New code should use
    /// <see cref="ExportProfile"/> so only Excel profiles are representable.
    /// </summary>
    public OfficeHtmlConversionProfile Profile {
        get => ExportProfile == ExcelHtmlExportProfile.VisualReview
            ? OfficeHtmlConversionProfile.ExcelVisualReview
            : OfficeHtmlConversionProfile.ExcelSemanticTables;
        set => ExportProfile = value switch {
            OfficeHtmlConversionProfile.ExcelSemanticTables => ExcelHtmlExportProfile.SemanticTables,
            OfficeHtmlConversionProfile.ExcelVisualReview => ExcelHtmlExportProfile.VisualReview,
            _ => throw new ArgumentOutOfRangeException(nameof(value), value, "The selected HTML conversion profile is not an Excel profile.")
        };
    }

    /// <summary>Shared engine profile used by the selected Excel export lane.</summary>
    public HtmlConversionProfile SharedProfile => ExportProfile == ExcelHtmlExportProfile.VisualReview
        ? HtmlConversionProfile.PositionedReview
        : HtmlConversionProfile.Semantic;

    /// <summary>Creates semantic worksheet-table export settings.</summary>
    public static ExcelHtmlSaveOptions CreateSemanticTablesProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.WordLike) => new() {
        ExportProfile = ExcelHtmlExportProfile.SemanticTables,
        Theme = theme
    };

    /// <summary>Creates positioned visual-review export settings.</summary>
    public static ExcelHtmlSaveOptions CreateVisualReviewProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.Report) => new() {
        ExportProfile = ExcelHtmlExportProfile.VisualReview,
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

    /// <summary>Maximum number of used-range rows exported per worksheet. Null uses the bounded default.</summary>
    public int? MaxRowsPerSheet { get; set; } = DefaultMaxRowsPerSheet;

    /// <summary>Maximum number of used-range columns exported per worksheet. Null uses the bounded default.</summary>
    public int? MaxColumnsPerSheet { get; set; } = DefaultMaxColumnsPerSheet;

    /// <summary>Maximum number of worksheet cells visited per semantic HTML table.</summary>
    public int MaxCellsPerSheet { get; set; } = DefaultMaxCellsPerSheet;

    /// <summary>Maximum number of merged-range records inspected per semantic HTML table.</summary>
    public int MaxMergedRangesPerSheet { get; set; } = DefaultMaxMergedRangesPerSheet;

    /// <summary>Text used for empty cells.</summary>
    public string EmptyCellText { get; set; } = string.Empty;

    /// <summary>
    /// Controls worksheet header semantics. Defaults to <see cref="ExcelHtmlHeaderMode.FirstRow"/>
    /// for compatibility with earlier OfficeIMO HTML output.
    /// </summary>
    public ExcelHtmlHeaderMode HeaderMode { get; set; } = ExcelHtmlHeaderMode.FirstRow;

    /// <summary>Options used by the existing Excel SVG visual export lane.</summary>
    public ExcelWorkbookImageExportOptions? VisualOptions { get; set; }

    /// <summary>
    /// Includes a deterministic semantic inventory for pivot tables. Interactive pivot behavior,
    /// caches, slicers, and timelines remain native workbook features and are not executed in HTML.
    /// </summary>
    public bool IncludePivotInventory { get; set; } = true;

    /// <summary>Creates an independent settings snapshot for one export operation.</summary>
    public ExcelHtmlSaveOptions Clone() => new() {
        ExportProfile = ExportProfile,
        DocumentOutput = DocumentOutput.Clone(),
        MaxRowsPerSheet = MaxRowsPerSheet,
        MaxColumnsPerSheet = MaxColumnsPerSheet,
        MaxCellsPerSheet = MaxCellsPerSheet,
        MaxMergedRangesPerSheet = MaxMergedRangesPerSheet,
        EmptyCellText = EmptyCellText,
        HeaderMode = HeaderMode,
        VisualOptions = VisualOptions,
        IncludePivotInventory = IncludePivotInventory
    };

    internal void Validate() {
        if (!Enum.IsDefined(typeof(ExcelHtmlExportProfile), ExportProfile)) {
            throw new ArgumentOutOfRangeException(nameof(ExportProfile), ExportProfile, "Excel HTML export profile is not supported.");
        }
        DocumentOutput.Validate();
        if (MaxRowsPerSheet.HasValue && MaxRowsPerSheet.Value <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxRowsPerSheet), "Maximum rows per worksheet must be positive when configured.");
        }
        if (MaxColumnsPerSheet.HasValue && MaxColumnsPerSheet.Value <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxColumnsPerSheet), "Maximum columns per worksheet must be positive when configured.");
        }
        if (MaxCellsPerSheet <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxCellsPerSheet), "Maximum cells per worksheet must be positive.");
        }
        if (MaxMergedRangesPerSheet <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxMergedRangesPerSheet), "Maximum merged ranges per worksheet must be positive.");
        }
    }
}

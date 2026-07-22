using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

/// <summary>
/// Options for exporting Excel workbooks and worksheets to HTML.
/// </summary>
public sealed class ExcelHtmlSaveOptions {
    /// <summary>Default maximum worksheet rows projected to semantic HTML.</summary>
    public const int DefaultMaxRowsPerSheet = 10000;

    /// <summary>Default maximum worksheet columns projected to semantic HTML.</summary>
    public const int DefaultMaxColumnsPerSheet = 1024;

    /// <summary>Default maximum worksheet cells visited while projecting semantic HTML.</summary>
    public const int DefaultMaxCellsPerSheet = 1000000;

    /// <summary>Excel-to-HTML lane to export. Defaults to semantic worksheet tables.</summary>
    public OfficeHtmlConversionProfile Profile { get; set; } = OfficeHtmlConversionProfile.ExcelSemanticTables;

    /// <summary>Theme used by the shared OfficeIMO HTML shell.</summary>
    public OfficeVisualThemeKind Theme { get; set; } = OfficeVisualThemeKind.WordLike;

    /// <summary>When true, emits the shared OfficeIMO CSS shell.</summary>
    public bool IncludeDefaultStyles { get; set; } = true;

    /// <summary>Optional document title.</summary>
    public string? Title { get; set; }

    /// <summary>Maximum number of used-range rows exported per worksheet. Null uses the bounded default.</summary>
    public int? MaxRowsPerSheet { get; set; } = DefaultMaxRowsPerSheet;

    /// <summary>Maximum number of used-range columns exported per worksheet. Null uses the bounded default.</summary>
    public int? MaxColumnsPerSheet { get; set; } = DefaultMaxColumnsPerSheet;

    /// <summary>Maximum number of worksheet cells visited per semantic HTML table.</summary>
    public int MaxCellsPerSheet { get; set; } = DefaultMaxCellsPerSheet;

    /// <summary>Text used for empty cells.</summary>
    public string EmptyCellText { get; set; } = string.Empty;

    /// <summary>
    /// Controls worksheet header semantics. Defaults to <see cref="ExcelHtmlHeaderMode.FirstRow"/>
    /// for compatibility with earlier OfficeIMO HTML output.
    /// </summary>
    public ExcelHtmlHeaderMode HeaderMode { get; set; } = ExcelHtmlHeaderMode.FirstRow;

    /// <summary>Options used by the existing Excel SVG visual export lane.</summary>
    public ExcelWorkbookImageExportOptions? VisualOptions { get; set; }

    internal void Validate() {
        if (MaxRowsPerSheet.HasValue && MaxRowsPerSheet.Value <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxRowsPerSheet), "Maximum rows per worksheet must be positive when configured.");
        }
        if (MaxColumnsPerSheet.HasValue && MaxColumnsPerSheet.Value <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxColumnsPerSheet), "Maximum columns per worksheet must be positive when configured.");
        }
        if (MaxCellsPerSheet <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxCellsPerSheet), "Maximum cells per worksheet must be positive.");
        }
    }
}

using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

/// <summary>
/// Options for saving an Excel-to-HTML capability gallery scenario.
/// </summary>
public sealed class ExcelHtmlCapabilityGalleryOptions {
    /// <summary>Stable scenario identifier used for file names and manifests.</summary>
    public string ScenarioId { get; set; } = "excel-rich-workbook";

    /// <summary>Human-readable scenario title.</summary>
    public string Title { get; set; } = "Excel Rich Workbook";

    /// <summary>Theme applied to generated HTML artifacts.</summary>
    public OfficeVisualThemeKind Theme { get; set; } = OfficeVisualThemeKind.Report;

    /// <summary>Options used by the visual SVG export lane.</summary>
    public ExcelWorkbookImageExportOptions? VisualOptions { get; set; }
}

using PdfCore = OfficeIMO.Pdf;
using DrawingCore = OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.Pdf;

/// <summary>
/// Options controlling first-party OfficeIMO PowerPoint-to-PDF export.
/// </summary>
public sealed class PowerPointPdfSaveOptions {
    /// <summary>PDF creation options passed to the first-party PDF engine.</summary>
    public PdfCore.PdfOptions? PdfOptions { get; set; }

    /// <summary>When true, supported slide pictures are embedded through the shared PDF image pipeline. Defaults to true.</summary>
    public bool IncludePictures { get; set; } = true;

    /// <summary>When true, supported slide auto-shapes are rendered through shared drawing primitives. Defaults to true.</summary>
    public bool IncludeAutoShapes { get; set; } = true;

    /// <summary>When true, slide text boxes are rendered as styled canvas text boxes. Defaults to true.</summary>
    public bool IncludeTextBoxes { get; set; } = true;

    /// <summary>When true, supported slide backgrounds are rendered before slide shapes. Defaults to true.</summary>
    public bool IncludeSlideBackgrounds { get; set; } = true;

    /// <summary>When true, supported slide tables are rendered through shared PDF table primitives. Defaults to true.</summary>
    public bool IncludeTables { get; set; } = true;

    /// <summary>When true, supported slide charts are rendered through shared vector chart primitives. Defaults to true.</summary>
    public bool IncludeCharts { get; set; } = true;

    /// <summary>Optional shared chart style applied to supported slide chart snapshots.</summary>
    public DrawingCore.OfficeChartStyle? ChartStyle { get; set; }

    /// <summary>Optional shared chart layout applied to supported slide chart snapshots.</summary>
    public DrawingCore.OfficeChartLayout? ChartLayout { get; set; }

    /// <summary>Warnings recorded during the latest export.</summary>
    public IList<PowerPointPdfExportWarning> Warnings { get; } = new List<PowerPointPdfExportWarning>();

    internal void ResetExportState() {
        Warnings.Clear();
    }
}

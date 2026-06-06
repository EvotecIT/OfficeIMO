using PdfCore = OfficeIMO.Pdf;
using DrawingCore = OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.Pdf;

/// <summary>
/// Options controlling first-party OfficeIMO PowerPoint-to-PDF export.
/// </summary>
public sealed class PowerPointPdfSaveOptions {
    /// <summary>PDF creation options passed to the first-party PDF engine.</summary>
    public PdfCore.PdfOptions? PdfOptions { get; set; }

    /// <summary>Optional PowerPoint-style font family used as the first-party PDF default font.</summary>
    public string? FontFamily { get; set; }

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

    /// <summary>When true, slides marked hidden in PowerPoint are exported. Defaults to false.</summary>
    public bool IncludeHiddenSlides { get; set; }

    private DrawingCore.OfficeImageFit _pictureFit = DrawingCore.OfficeImageFit.Stretch;

    /// <summary>
    /// Image fit mode used for uncropped PowerPoint pictures. Defaults to Stretch to match authored PowerPoint picture frames.
    /// </summary>
    public DrawingCore.OfficeImageFit PictureFit {
        get => _pictureFit;
        set {
            if (value != DrawingCore.OfficeImageFit.Stretch &&
                value != DrawingCore.OfficeImageFit.Contain &&
                value != DrawingCore.OfficeImageFit.Cover) {
                throw new ArgumentException("PowerPoint picture fit must be Stretch, Contain, or Cover.", nameof(value));
            }

            _pictureFit = value;
        }
    }

    /// <summary>When true, warns when stretched uncropped pictures visibly change their original aspect ratio. Defaults to false.</summary>
    public bool WarnOnPictureAspectRatioDistortion { get; set; }

    /// <summary>Optional shared chart style applied to supported slide chart snapshots.</summary>
    public DrawingCore.OfficeChartStyle? ChartStyle { get; set; }

    /// <summary>Optional shared chart layout applied to supported slide chart snapshots.</summary>
    public DrawingCore.OfficeChartLayout? ChartLayout { get; set; }

    /// <summary>Warnings recorded during the latest export.</summary>
    public IList<PowerPointPdfExportWarning> Warnings { get; } = new List<PowerPointPdfExportWarning>();

    /// <summary>
    /// Shared conversion report populated alongside <see cref="Warnings"/> for wrapper-friendly diagnostics.
    /// The report is cleared at the start of each export.
    /// </summary>
    public PdfCore.PdfConversionReport ConversionReport { get; } = new PdfCore.PdfConversionReport();

    internal void ResetExportState() {
        Warnings.Clear();
        ConversionReport.Clear();
    }
}

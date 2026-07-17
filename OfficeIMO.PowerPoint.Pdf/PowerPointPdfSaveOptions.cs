using PdfCore = OfficeIMO.Pdf;
using DrawingCore = OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.Pdf;

/// <summary>Page composition used for PowerPoint PDF export.</summary>
public enum PowerPointPdfPageLayout {
    /// <summary>One full-bleed slide per PDF page.</summary>
    Slides,
    /// <summary>One slide thumbnail with its speaker notes per portrait page.</summary>
    NotesPages,
    /// <summary>Several slide thumbnails per landscape handout page.</summary>
    Handouts
}

/// <summary>
/// Options controlling first-party OfficeIMO PowerPoint-to-PDF export.
/// </summary>
public sealed class PowerPointPdfSaveOptions {
    private int _handoutSlidesPerPage = 6;
    private PdfCore.PdfResourcePolicy _resourcePolicy = PdfCore.PdfResourcePolicy.CreateDefault();
    /// <summary>PDF creation options passed to the first-party PDF engine.</summary>
    public PdfCore.PdfOptions? PdfOptions { get; set; }

    /// <summary>Optional PowerPoint-style font family used as the first-party PDF default font.</summary>
    public string? FontFamily { get; set; }

    /// <summary>Host-resource policy. Defaults to balanced conversion: system fonts and bounded in-source resources are allowed, while local and remote reads are denied.</summary>
    public PdfCore.PdfResourcePolicy ResourcePolicy {
        get => _resourcePolicy;
        set => _resourcePolicy = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>
    /// Built-in generated-text fallback groups applied by the PowerPoint PDF converter.
    /// </summary>
    public PdfCore.PdfTextFallbackFeatures TextFallbacks { get; set; } = PdfCore.PdfTextFallbackFeatures.Default;

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

    /// <summary>Slide, notes-page, or handout PDF composition.</summary>
    public PowerPointPdfPageLayout PageLayout { get; set; } = PowerPointPdfPageLayout.Slides;

    /// <summary>Number of thumbnails per handout page. Supported values are 1, 2, 3, 4, 6, and 9.</summary>
    public int HandoutSlidesPerPage {
        get => _handoutSlidesPerPage;
        set {
            if (value != 1 && value != 2 && value != 3 && value != 4 && value != 6 && value != 9) {
                throw new ArgumentOutOfRangeException(nameof(value),
                    "Handout slides per page must be 1, 2, 3, 4, 6, or 9.");
            }
            _handoutSlidesPerPage = value;
        }
    }

    /// <summary>When true, notes-page and handout layouts include existing speaker-note text.</summary>
    public bool IncludeSpeakerNotes { get; set; } = true;

    /// <summary>Maximum nested group-shape depth rendered during PDF export. Defaults to 32.</summary>
    public int MaxGroupShapeDepth { get; set; } = 32;

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

    /// <summary>
    /// Applies a high-level export profile by setting the PowerPoint PDF options that correspond to that profile.
    /// </summary>
    public PowerPointPdfSaveOptions UseProfile(PdfCore.PdfExportProfile profile) {
        switch (profile) {
            case PdfCore.PdfExportProfile.Faithful:
                IncludePictures = true;
                IncludeAutoShapes = true;
                IncludeTextBoxes = true;
                IncludeSlideBackgrounds = true;
                IncludeTables = true;
                IncludeCharts = true;
                IncludeHiddenSlides = false;
                break;
            case PdfCore.PdfExportProfile.Lightweight:
                IncludePictures = false;
                IncludeAutoShapes = true;
                IncludeTextBoxes = true;
                IncludeSlideBackgrounds = false;
                IncludeTables = true;
                IncludeCharts = false;
                IncludeHiddenSlides = false;
                break;
            case PdfCore.PdfExportProfile.PrintReady:
                IncludePictures = true;
                IncludeAutoShapes = true;
                IncludeTextBoxes = true;
                IncludeSlideBackgrounds = true;
                IncludeTables = true;
                IncludeCharts = true;
                IncludeHiddenSlides = false;
                break;
            case PdfCore.PdfExportProfile.TextOnly:
                IncludePictures = false;
                IncludeAutoShapes = false;
                IncludeTextBoxes = true;
                IncludeSlideBackgrounds = false;
                IncludeTables = true;
                IncludeCharts = false;
                IncludeHiddenSlides = false;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unsupported PDF export profile.");
        }

        return this;
    }

    /// <summary>Warnings recorded during the latest export.</summary>
    internal IList<PowerPointPdfExportWarning> Warnings { get; private set; } = new List<PowerPointPdfExportWarning>();

    /// <summary>
    /// Shared conversion report populated alongside <see cref="Warnings"/> for wrapper-friendly diagnostics.
    /// The report is cleared at the start of each export.
    /// </summary>
    internal PdfCore.PdfConversionReport Report { get; private set; } = new PdfCore.PdfConversionReport();

    internal PowerPointPdfSaveOptions CloneForConversion() {
        var clone = (PowerPointPdfSaveOptions)MemberwiseClone();
        clone.ResourcePolicy = ResourcePolicy.Clone();
        clone.Warnings = new List<PowerPointPdfExportWarning>();
        clone.Report = new PdfCore.PdfConversionReport();
        return clone;
    }
}

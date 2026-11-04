using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates native PowerPoint-to-PDF options for Convert-OfficeDocument.</summary>
/// <example><summary>Create a three-slide handout.</summary>
/// <code>$options = New-OfficePowerPointPdfOptions -PageLayout Handouts -HandoutSlidesPerPage 3 -IncludeHiddenSlides
/// Convert-OfficeDocument -Path ./Briefing.pptx -To Pdf -OutputPath ./Briefing.pdf -Options $options</code></example>
/// <example><summary>Omit charts and SmartArt from a review copy.</summary>
/// <code>$options = New-OfficePowerPointPdfOptions -NoCharts -NoSmartArt -Font Aptos
/// Convert-OfficeDocument -Path ./Briefing.pptx -To Pdf -OutputPath ./Review.pdf -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficePowerPointPdfOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(PowerPointPdfExportOptions))]
public sealed class NewOfficePowerPointPdfOptionsCommand : OfficeCmdlet {
    /// <summary>Underlying low-level OfficeIMO PDF options.</summary>
    [Parameter]
    public OfficeIMO.Pdf.PdfCreateOptions? PdfOptions { get; set; }

    /// <summary>Default font family used when the presentation does not specify one.</summary>
    [Parameter]
    public string? Font { get; set; }

    /// <summary>Omit pictures.</summary>
    [Parameter]
    public SwitchParameter NoPictures { get; set; }

    /// <summary>Omit automatic shapes.</summary>
    [Parameter]
    public SwitchParameter NoAutoShapes { get; set; }

    /// <summary>Omit text boxes.</summary>
    [Parameter]
    public SwitchParameter NoTextBoxes { get; set; }

    /// <summary>Omit slide backgrounds.</summary>
    [Parameter]
    public SwitchParameter NoSlideBackgrounds { get; set; }

    /// <summary>Omit tables.</summary>
    [Parameter]
    public SwitchParameter NoTables { get; set; }

    /// <summary>Omit charts.</summary>
    [Parameter]
    public SwitchParameter NoCharts { get; set; }

    /// <summary>Omit SmartArt.</summary>
    [Parameter]
    public SwitchParameter NoSmartArt { get; set; }

    /// <summary>Include slides marked hidden.</summary>
    [Parameter]
    public SwitchParameter IncludeHiddenSlides { get; set; }

    /// <summary>PDF page layout, such as slides, notes, or handouts.</summary>
    [Parameter]
    public PowerPointPdfPageLayout? PageLayout { get; set; }

    /// <summary>Number of slides on each handout page.</summary>
    [Parameter]
    [ValidateRange(1, 9)]
    public int? HandoutSlidesPerPage { get; set; }

    /// <summary>Omit speaker notes.</summary>
    [Parameter]
    public SwitchParameter NoSpeakerNotes { get; set; }

    /// <summary>Maximum nested group-shape depth to render.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxGroupShapeDepth { get; set; }

    /// <summary>How pictures fit their shape bounds.</summary>
    [Parameter]
    public OfficeImageFit? PictureFit { get; set; }

    /// <summary>Report pictures whose requested fit distorts their aspect ratio.</summary>
    [Parameter]
    public SwitchParameter WarnOnPictureAspectRatioDistortion { get; set; }

    /// <summary>Chart visual style override.</summary>
    [Parameter]
    public OfficeChartStyle? ChartStyle { get; set; }

    /// <summary>Chart layout override.</summary>
    [Parameter]
    public OfficeChartLayout? ChartLayout { get; set; }

    /// <summary>Allow embedding fonts discovered on the current system.</summary>
    [Parameter]
    public SwitchParameter AllowSystemFontEmbedding { get; set; }

    /// <summary>Allow embedding installed fonts selected by a trusted presentation. System-font embedding must also be enabled.</summary>
    [Parameter]
    public SwitchParameter AllowDocumentFontEmbedding { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PowerPointPdfExportOptions();
        if (PdfOptions != null) options.PdfOptions = PdfOptions;
        if (!string.IsNullOrWhiteSpace(Font)) options.FontFamily = Font;
        options.IncludePictures = FeatureEnabled("Pictures");
        options.IncludeAutoShapes = FeatureEnabled("AutoShapes");
        options.IncludeTextBoxes = FeatureEnabled("TextBoxes");
        options.IncludeSlideBackgrounds = FeatureEnabled("SlideBackgrounds");
        options.IncludeTables = FeatureEnabled("Tables");
        options.IncludeCharts = FeatureEnabled("Charts");
        options.IncludeSmartArt = FeatureEnabled("SmartArt");
        if (BoundSwitch(nameof(IncludeHiddenSlides)) is bool hidden) options.IncludeHiddenSlides = hidden;
        if (PageLayout.HasValue) options.PageLayout = PageLayout.Value;
        if (HandoutSlidesPerPage.HasValue) options.HandoutSlidesPerPage = HandoutSlidesPerPage.Value;
        options.IncludeSpeakerNotes = FeatureEnabled("SpeakerNotes");
        if (MaxGroupShapeDepth.HasValue) options.MaxGroupShapeDepth = MaxGroupShapeDepth.Value;
        if (PictureFit.HasValue) options.PictureFit = PictureFit.Value;
        if (BoundSwitch(nameof(WarnOnPictureAspectRatioDistortion)) is bool warn) options.WarnOnPictureAspectRatioDistortion = warn;
        if (ChartStyle != null) options.ChartStyle = ChartStyle;
        if (ChartLayout != null) options.ChartLayout = ChartLayout;
        if (BoundSwitch(nameof(AllowSystemFontEmbedding)) is bool systemFonts) options.ResourcePolicy.AllowSystemFontEmbedding = systemFonts;
        if (BoundSwitch(nameof(AllowDocumentFontEmbedding)) is bool documentFonts) options.ResourcePolicy.AllowDocumentFontEmbedding = documentFonts;
        WriteObject(options);
    }

}

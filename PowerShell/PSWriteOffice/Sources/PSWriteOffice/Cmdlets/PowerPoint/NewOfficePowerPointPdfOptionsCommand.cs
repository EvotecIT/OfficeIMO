using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.Pdf;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates discoverable PowerPoint-to-PDF conversion options for Export-OfficeDocumentPdf.</summary>
/// <example>
///   <summary>Create a handout PDF with notes and hidden slides.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficePowerPointPdfOptions -PageLayout Handouts -HandoutSlidesPerPage 3 -IncludeSpeakerNotes -IncludeHiddenSlides
/// Export-OfficeDocumentPdf -InputPath .\Briefing.pptx -Path .\Briefing.pdf -PowerPointOptions $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePowerPointPdfOptions")]
[OutputType(typeof(PowerPointToPdfOptions))]
public sealed class NewOfficePowerPointPdfOptionsCommand : PSCmdlet {
    /// <summary>Underlying low-level OfficeIMO PDF options.</summary>
    [Parameter]
    public OfficeIMO.Pdf.PdfOptions? PdfOptions { get; set; }

    /// <summary>Default font family used when the presentation does not specify one.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>Render pictures.</summary>
    [Parameter]
    public SwitchParameter IncludePictures { get; set; }

    /// <summary>Render automatic shapes.</summary>
    [Parameter]
    public SwitchParameter IncludeAutoShapes { get; set; }

    /// <summary>Render text boxes.</summary>
    [Parameter]
    public SwitchParameter IncludeTextBoxes { get; set; }

    /// <summary>Render slide backgrounds.</summary>
    [Parameter]
    public SwitchParameter IncludeSlideBackgrounds { get; set; }

    /// <summary>Render tables.</summary>
    [Parameter]
    public SwitchParameter IncludeTables { get; set; }

    /// <summary>Render charts.</summary>
    [Parameter]
    public SwitchParameter IncludeCharts { get; set; }

    /// <summary>Render SmartArt.</summary>
    [Parameter]
    public SwitchParameter IncludeSmartArt { get; set; }

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

    /// <summary>Include speaker notes.</summary>
    [Parameter]
    public SwitchParameter IncludeSpeakerNotes { get; set; }

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
        var options = new PowerPointToPdfOptions();
        if (PdfOptions != null) options.PdfOptions = PdfOptions;
        if (!string.IsNullOrWhiteSpace(FontFamily)) options.FontFamily = FontFamily;
        SetBoundSwitch(nameof(IncludePictures), IncludePictures, value => options.IncludePictures = value);
        SetBoundSwitch(nameof(IncludeAutoShapes), IncludeAutoShapes, value => options.IncludeAutoShapes = value);
        SetBoundSwitch(nameof(IncludeTextBoxes), IncludeTextBoxes, value => options.IncludeTextBoxes = value);
        SetBoundSwitch(nameof(IncludeSlideBackgrounds), IncludeSlideBackgrounds, value => options.IncludeSlideBackgrounds = value);
        SetBoundSwitch(nameof(IncludeTables), IncludeTables, value => options.IncludeTables = value);
        SetBoundSwitch(nameof(IncludeCharts), IncludeCharts, value => options.IncludeCharts = value);
        SetBoundSwitch(nameof(IncludeSmartArt), IncludeSmartArt, value => options.IncludeSmartArt = value);
        SetBoundSwitch(nameof(IncludeHiddenSlides), IncludeHiddenSlides, value => options.IncludeHiddenSlides = value);
        if (PageLayout.HasValue) options.PageLayout = PageLayout.Value;
        if (HandoutSlidesPerPage.HasValue) options.HandoutSlidesPerPage = HandoutSlidesPerPage.Value;
        SetBoundSwitch(nameof(IncludeSpeakerNotes), IncludeSpeakerNotes, value => options.IncludeSpeakerNotes = value);
        if (MaxGroupShapeDepth.HasValue) options.MaxGroupShapeDepth = MaxGroupShapeDepth.Value;
        if (PictureFit.HasValue) options.PictureFit = PictureFit.Value;
        SetBoundSwitch(nameof(WarnOnPictureAspectRatioDistortion), WarnOnPictureAspectRatioDistortion, value => options.WarnOnPictureAspectRatioDistortion = value);
        if (ChartStyle != null) options.ChartStyle = ChartStyle;
        if (ChartLayout != null) options.ChartLayout = ChartLayout;
        SetBoundSwitch(nameof(AllowSystemFontEmbedding), AllowSystemFontEmbedding, value => options.ResourcePolicy.AllowSystemFontEmbedding = value);
        SetBoundSwitch(nameof(AllowDocumentFontEmbedding), AllowDocumentFontEmbedding, value => options.ResourcePolicy.AllowDocumentFontEmbedding = value);
        WriteObject(options);
    }

    private void SetBoundSwitch(string name, SwitchParameter value, System.Action<bool> setter) {
        if (MyInvocation.BoundParameters.ContainsKey(name)) setter(value.IsPresent);
    }
}

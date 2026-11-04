using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates native PDF layout-overlay layers, colors and resource limits.</summary>
/// <example><summary>Show paragraph regions and reading order without word or line boxes.</summary>
/// <code>$options = New-OfficePdfLayoutOverlayOptions -NoWords -NoLines -RegionColor '#FF00FF'
/// Export-OfficePdfLayoutOverlay -Path ./Report.pdf -OutputPath ./Order.svg -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficePdfLayoutOverlayOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(PdfLayoutDebugOverlayOptions))]
public sealed class NewOfficePdfLayoutOverlayOptionsCommand : OfficeCmdlet {
    /// <summary>Hide word boxes.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoWords { get; set; }
    /// <summary>Hide line boxes.</summary>
    [Parameter]
    public SwitchParameter NoLines { get; set; }
    /// <summary>Hide paragraph region boxes.</summary>
    [Parameter]
    public SwitchParameter NoRegions { get; set; }
    /// <summary>Hide reading-order labels.</summary>
    [Parameter]
    public SwitchParameter NoReadingOrder { get; set; }
    /// <summary>Word-box color or theme color role.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor? WordColor { get; set; }
    /// <summary>Line-box color or theme color role.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor? LineColor { get; set; }
    /// <summary>Paragraph-region color or theme color role.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor? RegionColor { get; set; }
    /// <summary>Maximum boxes and labels in one overlay.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaxElements { get; set; }
    /// <summary>Maximum allocated pixels in a PNG overlay.</summary>
    [Parameter, ValidateRange(1L, long.MaxValue)]
    public long? MaximumRasterPixelCount { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PdfLayoutDebugOverlayOptions {
            ShowWords = !NoWords.IsPresent, ShowLines = !NoLines.IsPresent,
            ShowRegions = !NoRegions.IsPresent, ShowReadingOrder = !NoReadingOrder.IsPresent
        };
        if (WordColor.HasValue) options.WordColor = WordColor.Value;
        if (LineColor.HasValue) options.LineColor = LineColor.Value;
        if (RegionColor.HasValue) options.RegionColor = RegionColor.Value;
        if (MaxElements.HasValue) options.MaxElements = MaxElements.Value;
        if (MaximumRasterPixelCount.HasValue) options.MaxRasterPixels = MaximumRasterPixelCount.Value;
        WriteObject(options);
    }
}

using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates discoverable rendering and tolerance settings for Compare-OfficePdfVisual.</summary>
/// <example>
///   <summary>Compare PDFs with a small rendering tolerance.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficePdfVisualComparisonOptions -ChannelTolerance 2 -AllowedDifferenceRatio 0.001 -MaxPages 50
/// Compare-OfficePdfVisual -ReferencePath .\Expected.pdf -DifferencePath .\Actual.pdf -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePdfVisualComparisonOptions")]
[OutputType(typeof(PdfVisualComparisonOptions))]
public sealed class NewOfficePdfVisualComparisonOptionsCommand : PSCmdlet {
    /// <summary>Render scale applied before comparison.</summary>
    [Parameter] [ValidateRange(double.Epsilon, double.MaxValue)] public double? Scale { get; set; }
    /// <summary>Maximum per-channel byte difference treated as equal.</summary>
    [Parameter] public byte? ChannelTolerance { get; set; }
    /// <summary>Maximum differing-pixel ratio treated as equal.</summary>
    [Parameter] [ValidateRange(0d, 1d)] public double? AllowedDifferenceRatio { get; set; }
    /// <summary>Page alignment used for differently sized renders.</summary>
    [Parameter] public PdfVisualPageAlignment? Alignment { get; set; }
    /// <summary>Background color name or hex value.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? BackgroundColor { get; set; }
    /// <summary>Maximum pages compared.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxPages { get; set; }
    /// <summary>Maximum pixels accepted per rendered image.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaxPixelsPerImage { get; set; }
    /// <summary>Maximum pixels accepted across the comparison.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaxTotalPixels { get; set; }
    /// <summary>Maximum total bytes retained for comparison artifacts.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaxTotalOutputBytes { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PdfVisualComparisonOptions();
        if (Scale.HasValue) options.Scale = Scale.Value;
        if (ChannelTolerance.HasValue) options.ChannelTolerance = ChannelTolerance.Value;
        if (AllowedDifferenceRatio.HasValue) options.AllowedDifferenceRatio = AllowedDifferenceRatio.Value;
        if (Alignment.HasValue) options.Alignment = Alignment.Value;
        if (!string.IsNullOrWhiteSpace(BackgroundColor)) options.Background = OfficeColor.Parse(BackgroundColor!);
        if (MaxPages.HasValue) options.MaxPages = MaxPages.Value;
        if (MaxPixelsPerImage.HasValue) options.MaxPixelsPerImage = MaxPixelsPerImage.Value;
        if (MaxTotalPixels.HasValue) options.MaxTotalPixels = MaxTotalPixels.Value;
        if (MaxTotalOutputBytes.HasValue) options.MaxTotalOutputBytes = MaxTotalOutputBytes.Value;
        WriteObject(options);
    }
}

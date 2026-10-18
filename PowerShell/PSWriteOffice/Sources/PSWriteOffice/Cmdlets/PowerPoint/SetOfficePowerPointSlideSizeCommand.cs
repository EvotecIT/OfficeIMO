using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets the slide size for a PowerPoint presentation.</summary>
/// <para>Supports common presets as well as explicit width and height in centimeters, inches, points, or EMUs.</para>
/// <example>
///   <summary>Set a standard widescreen presentation size.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\Examples\Documents\PowerPointWidescreen.pptx {
///     Set-OfficePowerPointSlideSize -Preset Screen16x9
///     Add-OfficePowerPointSlide -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Widescreen deck'
/// }</code>
///   <para>Applies the 16:9 widescreen preset before adding slides.</para>
/// </example>
/// <example>
///   <summary>Set a custom size in centimeters.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = New-OfficePowerPoint -Path .\Examples\Documents\PowerPointCustomSize.pptx -NoSave
/// Set-OfficePowerPointSlideSize -Presentation $ppt -WidthCm 25.4 -HeightCm 14.0
/// Add-OfficePowerPointSlide -Presentation $ppt -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Custom size'
/// $ppt | Close-OfficePowerPoint -Save</code>
///   <para>Sets the presentation slide size to a custom 25.4 x 14.0 cm layout.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointSlideSize", DefaultParameterSetName = ParameterSetPreset)]
[Alias("PptSlideSize")]
[OutputType(typeof(PowerPointSlideSize))]
public sealed class SetOfficePowerPointSlideSizeCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetPreset = "Preset";
    private const string ParameterSetCentimeters = "Centimeters";
    private const string ParameterSetInches = "Inches";
    private const string ParameterSetPoints = "Points";
    private const string ParameterSetEmus = "Emus";

    /// <summary>Presentation to update (optional inside New-OfficePowerPoint).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Preset slide size to apply.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPreset)]
    public PowerPointSlideSizePreset Preset { get; set; }

    /// <summary>Apply the preset in portrait orientation.</summary>
    [Parameter(ParameterSetName = ParameterSetPreset)]
    public SwitchParameter Portrait { get; set; }

    /// <summary>Custom slide width in centimeters.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetCentimeters)]
    public double WidthCm { get; set; }

    /// <summary>Custom slide height in centimeters.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetCentimeters)]
    public double HeightCm { get; set; }

    /// <summary>Custom slide width in inches.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInches)]
    public double WidthInches { get; set; }

    /// <summary>Custom slide height in inches.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInches)]
    public double HeightInches { get; set; }

    /// <summary>Custom slide width in points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPoints)]
    public double WidthPoints { get; set; }

    /// <summary>Custom slide height in points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPoints)]
    public double HeightPoints { get; set; }

    /// <summary>Custom slide width in EMUs.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetEmus)]
    public long WidthEmus { get; set; }

    /// <summary>Custom slide height in EMUs.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetEmus)]
    public long HeightEmus { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var presentation = Presentation ?? PowerPointDslContext.Current?.Presentation
                ?? throw new InvalidOperationException("Presentation was not provided. Use -Presentation or run inside New-OfficePowerPoint.");

            var slideSize = presentation.SlideSize;
            switch (ParameterSetName) {
                case ParameterSetPreset:
                    slideSize.SetPreset(Preset, Portrait.IsPresent);
                    break;
                case ParameterSetCentimeters:
                    slideSize.SetSizeCm(WidthCm, HeightCm);
                    break;
                case ParameterSetInches:
                    slideSize.SetSizeInches(WidthInches, HeightInches);
                    break;
                case ParameterSetPoints:
                    slideSize.SetSizePoints(WidthPoints, HeightPoints);
                    break;
                case ParameterSetEmus:
                    slideSize.SetSizeEmus(WidthEmus, HeightEmus);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported parameter set '{ParameterSetName}'.");
            }

            WritePassThru(slideSize);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointSetSlideSizeFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}

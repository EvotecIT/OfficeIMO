using System.Management.Automation;
using OfficeIMO.ChartForgeX;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a ChartForgeX artifact, portable SVG, or converted Office visual to a PowerPoint slide.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointVisual")]
[Alias("PptVisual")]
[OutputType(typeof(PowerPointPicture))]
public sealed class AddOfficePowerPointVisualCommand : OfficeVisualCommandBase {
    /// <summary>ChartForgeX VisualArtifact, OfficeVisualSource, OfficeVisualConversionResult, or SVG file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public object InputObject { get; set; } = null!;

    /// <summary>Target slide. Inside the PowerPoint DSL, the current slide is used by default.</summary>
    [Parameter]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Left offset in points.</summary>
    [Parameter]
    public double X { get; set; }

    /// <summary>Top offset in points.</summary>
    [Parameter]
    public double Y { get; set; }

    /// <summary>Emit the picture added to the slide.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        PowerPointSlide slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();
        var picture = slide.AddVisualArtifact(ResolveVisual(InputObject), X, Y);
        if (PassThru.IsPresent) {
            WriteObject(picture);
        }
    }
}
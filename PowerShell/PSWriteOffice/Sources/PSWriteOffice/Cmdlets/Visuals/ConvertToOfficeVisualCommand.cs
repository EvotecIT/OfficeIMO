using System.Management.Automation;
using OfficeIMO.ChartForgeX;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.Visuals;

/// <summary>Converts a ChartForgeX visual artifact into a reusable OfficeIMO representation.</summary>
/// <example>
///   <summary>Prepare one chart for multiple Office outputs.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$officeVisual = $artifact | ConvertTo-OfficeVisual -Width 420 -SvgPolicy RasterizeWhenNeeded</code>
///   <para>Creates one fidelity report and placement payload that can be reused in Word, Excel, PowerPoint, and PDF.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeVisual")]
[Alias("OfficeVisual")]
[OutputType(typeof(OfficeVisualConversionResult))]
public sealed class ConvertToOfficeVisualCommand : OfficeVisualCommandBase
{
    /// <summary>ChartForgeX VisualArtifact, OfficeVisualSource, prepared conversion, or portable SVG file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public object InputObject { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(ResolveVisual(InputObject));
}

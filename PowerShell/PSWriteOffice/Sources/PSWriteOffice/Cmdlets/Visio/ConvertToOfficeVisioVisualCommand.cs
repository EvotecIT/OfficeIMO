using System.Management.Automation;
using OfficeIMO.ChartForgeX;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Converts CFX semantic visual-artifact input into a native editable OfficeIMO.Visio document.</summary>
/// <example>
///   <summary>Convert an ImagePlayground topology into editable Visio objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$visio = $topology | ConvertTo-ImageVisualArtifact | ConvertTo-OfficeVisioVisual
/// $visio.Document | Save-OfficeVisio -Path .\Topology.vsdx</code>
///   <para>Consumes the portable CFX semantic payload and returns the document, page, and fidelity report.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeVisioVisual")]
[OutputType(typeof(OfficeVisioVisualConversionResult))]
public sealed class ConvertToOfficeVisioVisualCommand : OfficeVisioVisualCommandBase
{
    /// <summary>Typed CFX artifact, semantic envelope/JSON, ImagePlayground portable artifact, or prior conversion result.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public object InputObject { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (BufferPipelineByte(InputObject))
        {
            return;
        }

        WriteObject(ResolveVisioVisual(InputObject));
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        byte[]? jsonBytes = CompletePipelineBytes();
        if (jsonBytes != null)
        {
            WriteObject(ResolveVisioVisual(jsonBytes));
        }
    }
}

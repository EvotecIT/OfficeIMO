using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Loads an existing .vsdx file as an OfficeIMO.Visio document.</summary>
/// <example>
///   <summary>Open and inspect a Visio document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$diagram = Get-OfficeVisio -Path .\ServiceMap.vsdx
/// Get-OfficeVisioInfo -Document $diagram -AsText</code>
///   <para>Loads an existing .vsdx file and creates a deterministic inspection snapshot.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeVisio")]
[Alias("VisioOpen")]
[OutputType(typeof(VisioDocument))]
public sealed class GetOfficeVisioCommand : PSCmdlet
{
    /// <summary>Visio .vsdx path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WriteObject(VisioDocument.Load(VisioCommandUtilities.ResolvePath(this, Path)));
    }
}

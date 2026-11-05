using System.IO;
using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Loads an existing .vsdx file as an OfficeIMO.Visio document.</summary>
/// <example>
///   <summary>Open and inspect a Visio document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$diagram = Open-OfficeVisio -Path .\ServiceMap.vsdx
/// Open-OfficeVisioInfo -Document $diagram -AsText</code>
///   <para>Loads an existing .vsdx file and creates a deterministic inspection snapshot.</para>
/// </example>
[Cmdlet(VerbsCommon.Open, "OfficeVisio", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(VisioDocument))]
public sealed class OpenOfficeVisioCommand : OfficeLifecycleCmdlet<VisioDocument>
{
    /// <summary>An optional authoring block; omission returns the live document.</summary>
    [Parameter(ParameterSetName = "Path", Position = 1)]
    [Parameter(ParameterSetName = "LiteralPath", Position = 1)]
    [ValidateNotNull]
    public override ScriptBlock? Content { get => base.Content; set => base.Content = value; }

    /// <summary>Emit the FileInfo after a completed authoring block.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public override SwitchParameter PassThru { get => base.PassThru; set => base.PassThru = value; }

    /// <summary>Open the saved document after a completed authoring block.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public override SwitchParameter Show { get => base.Show; set => base.Show = value; }

    /// <summary>Visio .vsdx path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;
    /// <summary>Exact source path without wildcard expansion.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    public string LiteralPath { get; set; } = string.Empty;


    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var path in OfficePath.ResolveInput(SessionState, ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath"))
            RunDocument(path, false, false, source => VisioDocument.Load(source!), (document, destination, allowLoss) => { document.Save(destination); return null; });
    }
    /// <inheritdoc />
    protected override void InvokeContent(VisioDocument document) {
        using var context = VisioDslContext.Enter(document);
        using (context.Push(document.Pages[0])) Content!.InvokeReturnAsIs();
    }
}

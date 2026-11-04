using PSWriteOffice.Foundation;
using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Loads a project without calculating or executing its content.</summary>
/// <example>
///   <summary>Loads a project without calculating or executing its content.</summary>
///   <code>Open-OfficeProject -Path ./delivery.mpp -ReadOnly</code>
/// </example>
[Cmdlet(VerbsCommon.Open, "OfficeProject", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(ProjectDocument))]
public sealed class OpenOfficeProjectCommand : OfficeLifecycleCmdlet<ProjectDocument> {
    /// <summary>An optional authoring block; omission returns the live document.</summary>
    [Parameter(ParameterSetName = "Path", Position = 1)]
    [Parameter(ParameterSetName = "LiteralPath", Position = 1)]
    [ValidateNotNull]
    public override ScriptBlock? Content { get => base.Content; set => base.Content = value; }

    /// <summary>Emit the FileInfo after a completed authoring block.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public override SwitchParameter PassThru { get => base.PassThru; set => base.PassThru = value; }

    /// <summary>Explicitly accept native-save loss for an authoring block.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public override SwitchParameter AllowLoss { get => base.AllowLoss; set => base.AllowLoss = value; }

    /// <summary>Literal filesystem path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;
    /// <summary>Exact source path without wildcard expansion.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    public string LiteralPath { get; set; } = string.Empty;


    /// <summary>Open without allowing edits.</summary>
    [Parameter]
    public SwitchParameter ReadOnly { get; set; }

    /// <summary>Maximum source size in bytes.</summary>
    [Parameter]
    [ValidateRange(1L, int.MaxValue)]
    public long MaxInputBytes { get; set; } = 67108864;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var path in OfficePath.ResolveInput(SessionState, ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath"))
            RunDocument(path, false, false, source => ProjectDocument.Load(source!, new ProjectLoadOptions {
                AccessMode = ReadOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
                PersistenceMode = DocumentPersistenceMode.Explicit, MaxInputBytes = MaxInputBytes
            }), (document, destination, allowLoss) => { document.Save(destination, new ProjectSaveOptions { AllowLoss = allowLoss }); return null; });
    }
    /// <inheritdoc />
    protected override void InvokeContent(ProjectDocument document) {
        using (OfficeDocumentContext.Push(document)) Content!.InvokeReturnAsIs();
    }
}

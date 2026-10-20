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
///   <code>Get-OfficeProject -Path ./delivery.mpp -ReadOnly</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeProject")]
[OutputType(typeof(ProjectDocument))]
public sealed class GetOfficeProjectCommand : PSCmdlet {
    /// <summary>Literal filesystem path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "LiteralPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open without allowing edits.</summary>
    [Parameter]
    public SwitchParameter ReadOnly { get; set; }

    /// <summary>Maximum source size in bytes.</summary>
    [Parameter]
    [ValidateRange(1L, int.MaxValue)]
    public long MaxInputBytes { get; set; } = 67108864;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WriteObject(ProjectDocument.Load(OfficeFileCommandUtilities.ResolvePath(this, Path), new ProjectLoadOptions {
            AccessMode = ReadOnly.IsPresent ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit,
            MaxInputBytes = MaxInputBytes
        }));
    }
}

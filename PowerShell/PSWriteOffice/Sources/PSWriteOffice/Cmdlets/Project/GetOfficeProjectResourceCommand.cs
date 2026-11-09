using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Returns project resources.</summary>
/// <example>
///   <summary>Returns project resources.</summary>
///   <code>$project | Get-OfficeProjectResource</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeProjectResource", DefaultParameterSetName = "Document")]
[OutputType(typeof(ProjectResource))]
public sealed class GetOfficeProjectResourceCommand : PSCmdlet {
    /// <summary>Project document.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document")]
    public ProjectDocument Document { get; set; } = null!;
    /// <summary>Optional stable identities.</summary>
    [Parameter]
    public int[]? Uid { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        Document ??= OfficeDocumentContext.Find<ProjectDocument>() ?? throw new PSInvalidOperationException("Provide -Document or use a Project content block.");
        foreach (var item in Document.Resources) {
            if (Uid != null && !Uid.Contains(item.Uid)) continue;
            WriteObject(item);
        }
    }
}

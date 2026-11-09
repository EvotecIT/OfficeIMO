using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Returns project assignments.</summary>
/// <example>
///   <summary>Returns project assignments.</summary>
///   <code>$project | Get-OfficeProjectAssignment</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeProjectAssignment", DefaultParameterSetName = "Document")]
[OutputType(typeof(ProjectAssignment))]
public sealed class GetOfficeProjectAssignmentCommand : PSCmdlet {
    /// <summary>Project document.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document")]
    public ProjectDocument Document { get; set; } = null!;
    /// <summary>Optional stable identities.</summary>
    [Parameter]
    public int[]? Uid { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        Document ??= OfficeDocumentContext.Find<ProjectDocument>() ?? throw new PSInvalidOperationException("Provide -Document or use a Project content block.");
        foreach (var item in Document.Assignments) {
            if (Uid != null && !Uid.Contains(item.Uid)) continue;
            WriteObject(item);
        }
    }
}

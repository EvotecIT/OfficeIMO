using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Captures a portable report from a current calculated schedule.</summary>
/// <example>
///   <summary>Captures a portable report from a current calculated schedule.</summary>
///   <code>$view = $project | Get-OfficeProjectView -Schedule $schedule -Kind Gantt</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeProjectView", DefaultParameterSetName = "Document")]
[OutputType(typeof(ProjectView))]
public sealed class GetOfficeProjectViewCommand : PSCmdlet {
    /// <summary>Project document.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;
    /// <summary>Current schedule result for this project revision. Omission calculates a report snapshot without changing the document.</summary>
    [Parameter]
    public ProjectScheduleResult? Schedule { get; set; }
    /// <summary>Report presentation.</summary>
    [Parameter(ParameterSetName = "Document")]
    public ProjectViewKind Kind { get; set; } = ProjectViewKind.Gantt;
    /// <summary>Optional explicit report settings.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Options")]
    public ProjectViewOptions? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        Document ??= OfficeDocumentContext.Find<ProjectDocument>() ?? throw new PSInvalidOperationException("Provide -Document or use a Project content block.");
        var schedule = Schedule ?? Document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        WriteObject(Document.CreateView(schedule, Options ?? new ProjectViewOptions { Kind = Kind }));
    }
}

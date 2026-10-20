using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Captures a portable report from a current calculated schedule.</summary>
/// <example>
///   <summary>Captures a portable report from a current calculated schedule.</summary>
///   <code>$view = $project | Get-OfficeProjectView -Schedule $schedule -Kind Gantt</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeProjectView")]
[OutputType(typeof(ProjectView))]
public sealed class GetOfficeProjectViewCommand : PSCmdlet {
    /// <summary>Project document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;
    /// <summary>Current schedule result for this project revision. Omission calculates a report snapshot without changing the document.</summary>
    [Parameter]
    public ProjectScheduleResult? Schedule { get; set; }
    /// <summary>Report presentation.</summary>
    [Parameter]
    public ProjectViewKind Kind { get; set; } = ProjectViewKind.Gantt;
    /// <summary>Optional explicit report settings.</summary>
    [Parameter]
    public ProjectViewOptions? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Options != null && MyInvocation.BoundParameters.ContainsKey(nameof(Kind)))
            throw new PSArgumentException("Choose Options or Kind.");
        var schedule = Schedule ?? Document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        WriteObject(Document.CreateView(schedule, Options ?? new ProjectViewOptions { Kind = Kind }));
    }
}

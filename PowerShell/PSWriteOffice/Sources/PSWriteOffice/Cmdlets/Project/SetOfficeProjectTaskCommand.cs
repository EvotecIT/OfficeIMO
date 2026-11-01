using System;
using System.Management.Automation;
using OfficeIMO.Project;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Updates selected task fields without implicitly recalculating its schedule.</summary>
/// <example>
///   <summary>Change task progress and its working duration.</summary>
///   <code>$project | Get-OfficeProjectTask -Uid 7 | Set-OfficeProjectTask -DurationDays 4 -PercentComplete 25</code>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeProjectTask", DefaultParameterSetName = "Task")]
[OutputType(typeof(ProjectTask))]
public sealed class SetOfficeProjectTaskCommand : PSCmdlet {
    /// <summary>Editable task returned by Get-OfficeProjectTask.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectTask Task { get; set; } = null!;

    /// <summary>New task name.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Working duration in days.</summary>
    [Parameter]
    [ValidateRange(0d, 1000000d)]
    public decimal? DurationDays { get; set; }

    /// <summary>Stored local start date.</summary>
    [Parameter]
    public DateTime? Start { get; set; }

    /// <summary>Stored local finish date.</summary>
    [Parameter]
    public DateTime? Finish { get; set; }

    /// <summary>Duration completion from zero to one hundred.</summary>
    [Parameter]
    [ValidateRange(0, 100)]
    public int? PercentComplete { get; set; }

    /// <summary>Explicit manual scheduling choice.</summary>
    [Parameter]
    public bool? Manual { get; set; }

    /// <summary>Return the changed task.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Name != null) Task.Name = Name;
        if (DurationDays.HasValue) {
            Task.Duration = ProjectDuration.WorkingDays(DurationDays.Value);
            Task.IsMilestone = DurationDays == 0;
        }
        if (Start.HasValue) Task.Start = Start;
        if (Finish.HasValue) Task.Finish = Finish;
        if (PercentComplete.HasValue) Task.PercentComplete = PercentComplete;
        if (Manual.HasValue) Task.IsManual = Manual;
        if (PassThru.IsPresent) WriteObject(Task);
    }
}

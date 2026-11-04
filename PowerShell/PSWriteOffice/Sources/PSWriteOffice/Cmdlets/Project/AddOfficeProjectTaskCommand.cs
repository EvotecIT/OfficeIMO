using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Adds a task or summary with a duration expressed in working days.</summary>
/// <example>
///   <summary>Adds a task or summary with a duration expressed in working days.</summary>
///   <code>$project | Add-OfficeProjectTask -Name Design -DurationDays 3 -PassThru</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeProjectTask", DefaultParameterSetName = "Document")]
[OutputType(typeof(ProjectTask))]
public sealed class AddOfficeProjectTaskCommand : PSCmdlet {
    /// <summary>Project to inspect or change.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;

    /// <summary>Task name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Working duration in days.</summary>
    [Parameter]
    [ValidateRange(0d, 1000000d)]
    public decimal? DurationDays { get; set; }

    /// <summary>Parent summary UID; omission adds a root task.</summary>
    [Parameter]
    public int? ParentUid { get; set; }

    /// <summary>Create a summary task.</summary>
    [Parameter]
    public SwitchParameter Summary { get; set; }

    /// <summary>Use manual task scheduling instead of automatic scheduling.</summary>
    [Parameter]
    public SwitchParameter Manual { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var collection = ParentUid.HasValue ? Document.Tasks.GetByUid(ParentUid.Value).Children : Document.Tasks;
        var task = Summary.IsPresent ? collection.AddSummary(Name) : collection.Add(Name);
        task.IsManual = Manual.IsPresent;
        task.IsActive = true;
        task.IsNull = false;
        task.IsMilestone = DurationDays == 0;
        task.IsCritical = false;
        task.EffortDriven = false;
        if (DurationDays.HasValue) task.Duration = ProjectDuration.WorkingDays(DurationDays.Value);
        Document.NormalizeDisplayIds();
        if (PassThru.IsPresent) WriteObject(task);
    }
}

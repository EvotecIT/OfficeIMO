using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Explicitly calculates a schedule and optionally applies its dates.</summary>
/// <example>
///   <summary>Explicitly calculates a schedule and optionally applies its dates.</summary>
///   <code>$project | Invoke-OfficeProjectSchedule -CalculateAssignments -Apply</code>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeProjectSchedule", SupportsShouldProcess = true)]
[OutputType(typeof(ProjectScheduleResult))]
public sealed class InvokeOfficeProjectScheduleCommand : PSCmdlet {
    /// <summary>Project to inspect or change.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;

    /// <summary>Also calculate assignment work and costs.</summary>
    [Parameter]
    public SwitchParameter CalculateAssignments { get; set; }

    /// <summary>Apply an error-free result to this project. Otherwise return a calculation without changing stored dates.</summary>
    [Parameter]
    public SwitchParameter Apply { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Apply.IsPresent && !ShouldProcess(Document.Name ?? "Project", "Apply calculated schedule")) return;
        var result = Document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = CalculateAssignments.IsPresent });
        if (Apply.IsPresent) Document.ApplySchedule(result);
        WriteObject(result);
    }
}

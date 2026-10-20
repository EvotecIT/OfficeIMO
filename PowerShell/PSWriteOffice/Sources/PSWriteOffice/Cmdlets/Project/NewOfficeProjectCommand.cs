using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Creates a live project with an optional standard working calendar.</summary>
/// <example>
///   <summary>Creates a live project with an optional standard working calendar.</summary>
///   <code>$project = New-OfficeProject -Name Delivery -StartDate '2026-10-20 08:00'</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeProject")]
[OutputType(typeof(ProjectDocument))]
public sealed class NewOfficeProjectCommand : PSCmdlet {
    /// <summary>Project name.</summary>
    [Parameter(Position = 0)]
    public string? Name { get; set; }

    /// <summary>Explicit project start; omitted dates remain unspecified.</summary>
    [Parameter]
    public DateTime? StartDate { get; set; }

    /// <summary>Do not create a Monday-to-Friday working calendar.</summary>
    [Parameter]
    public SwitchParameter NoStandardCalendar { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var project = ProjectDocument.Create();
        project.Name = Name;
        if (StartDate.HasValue) {
            project.Settings.StartDate = StartDate;
            project.Settings.ScheduleFromStart = true;
        }
        if (!NoStandardCalendar.IsPresent) project.Calendar = project.Calendars.AddStandardWorkingWeek();
        WriteObject(project);
    }
}

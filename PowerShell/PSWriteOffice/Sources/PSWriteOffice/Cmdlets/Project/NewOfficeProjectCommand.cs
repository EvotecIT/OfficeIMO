using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Creates a live project with an optional standard working calendar.</summary>
/// <example>
///   <summary>Creates a live project with an optional standard working calendar.</summary>
///   <code>$project = New-OfficeProject -Name Delivery -StartDate '2026-10-20 08:00'</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeProject", DefaultParameterSetName = "Live", SupportsShouldProcess = true)]
[OutputType(typeof(ProjectDocument))]
public sealed class NewOfficeProjectCommand : OfficeLifecycleCmdlet<ProjectDocument> {
    /// <summary>Optional live association; required for a completed content block.</summary>
    [Parameter(Position = 0, ParameterSetName = "Live")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public string? Path { get; set; }
    /// <summary>Authoring block completed before atomic publication.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Content"), ValidateNotNull]
    public override ScriptBlock? Content { get => base.Content; set => base.Content = value; }
    /// <summary>Accept reductions reported by native save.</summary>
    [Parameter(ParameterSetName = "Content")]
    public override SwitchParameter AllowLoss { get => base.AllowLoss; set => base.AllowLoss = value; }
    /// <summary>Return the saved FileInfo after the block completes.</summary>
    [Parameter(ParameterSetName = "Content")]
    public override SwitchParameter PassThru { get => base.PassThru; set => base.PassThru = value; }
    /// <summary>Permit replacement of the associated destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Project name.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Explicit project start; omitted dates remain unspecified.</summary>
    [Parameter]
    public DateTime? StartDate { get; set; }

    /// <summary>Do not create a Monday-to-Friday working calendar.</summary>
    [Parameter]
    public SwitchParameter NoStandardCalendar { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        RunDocument(Path, true, Force, path => {
            var project = path == null ? ProjectDocument.Create() : ProjectDocument.Create(path);
            project.Name = Name;
            if (StartDate.HasValue) { project.Settings.StartDate = StartDate; project.Settings.ScheduleFromStart = true; }
            if (!NoStandardCalendar) project.Calendar = project.Calendars.AddStandardWorkingWeek();
            return project;
        }, (project, path, allowLoss) => {
            project.Save(path, new ProjectSaveOptions { AllowLoss = allowLoss, FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace });
            return null;
        });
    }
    /// <inheritdoc />
    protected override void InvokeContent(ProjectDocument document) {
        using (OfficeDocumentContext.Push(document)) Content!.InvokeReturnAsIs();
    }
}

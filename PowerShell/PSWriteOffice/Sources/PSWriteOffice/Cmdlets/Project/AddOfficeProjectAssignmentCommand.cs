using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Assigns a project resource to a task by stable UID.</summary>
/// <example>
///   <summary>Assigns a project resource to a task by stable UID.</summary>
///   <code>$project | Add-OfficeProjectAssignment -TaskUid 1 -ResourceUid 1 -UnitsPercent 50</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeProjectAssignment", DefaultParameterSetName = "Document")]
[OutputType(typeof(ProjectAssignment))]
public sealed class AddOfficeProjectAssignmentCommand : PSCmdlet {
    /// <summary>Project to inspect or change.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;

    /// <summary>Stable task UID.</summary>
    [Parameter(Mandatory = true)]
    public int TaskUid { get; set; }

    /// <summary>Stable resource UID.</summary>
    [Parameter(Mandatory = true)]
    public int ResourceUid { get; set; }

    /// <summary>Assignment units as a percentage.</summary>
    [Parameter]
    [ValidateRange(0d, 1000000d)]
    public decimal UnitsPercent { get; set; } = 100;

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var assignment = Document.Assignments.Add(Document.Tasks.GetByUid(TaskUid),
            Document.Resources.GetByUid(ResourceUid), ProjectUnits.Percent(UnitsPercent));
        if (PassThru.IsPresent) WriteObject(assignment);
    }
}

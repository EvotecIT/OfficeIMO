using PSWriteOffice.Services.Project;
using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Assigns a project resource to a task by stable UID.</summary>
/// <example>
///   <summary>Assigns a project resource to a task by stable UID.</summary>
///   <code>$project | Add-OfficeProjectAssignment -Task 1 -Resource 1 -UnitsPercent 50</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeProjectAssignment", DefaultParameterSetName = "Document")]
[OutputType(typeof(ProjectAssignment))]
[Alias("ProjectAssignment") ]
public sealed class AddOfficeProjectAssignmentCommand : PSCmdlet {
    /// <summary>Project to inspect or change.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document")]
    public ProjectDocument Document { get; set; } = null!;

    /// <summary>Stable task UID.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [OfficeProjectTaskTransformation]
    public OfficeProjectTaskSelector Task { get; set; } = null!;

    /// <summary>Stable resource UID.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [OfficeProjectResourceTransformation]
    public OfficeProjectResourceSelector Resource { get; set; } = null!;

    /// <summary>Assignment units as a percentage.</summary>
    [Parameter]
    [ValidateRange(0d, 1000000d)]
    public decimal UnitsPercent { get; set; } = 100;

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        Document ??= OfficeDocumentContext.Find<ProjectDocument>() ?? throw new PSInvalidOperationException("Provide -Document or use a Project content block.");
        var assignment = Document.Assignments.Add(Task.Resolve(Document),
            Resource.Resolve(Document), ProjectUnits.Percent(UnitsPercent));
        if (PassThru.IsPresent) WriteObject(assignment);
    }
}

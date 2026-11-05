using PSWriteOffice.Services.Project;
using PSWriteOffice.Foundation;
using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Adds a finish-to-start relationship between two tasks.</summary>
/// <example>
///   <summary>Adds a finish-to-start relationship between two tasks.</summary>
///   <code>$project | Add-OfficeProjectDependency -Predecessor 1 -Successor 2</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeProjectDependency", DefaultParameterSetName = "Document")]
[OutputType(typeof(ProjectDependency))]
[Alias("ProjectDependency") ]
public sealed class AddOfficeProjectDependencyCommand : PSCmdlet {
    /// <summary>Project to inspect or change.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document")]
    public ProjectDocument Document { get; set; } = null!;

    /// <summary>Predecessor task UID.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [OfficeProjectTaskTransformation]
    public OfficeProjectTaskSelector Predecessor { get; set; } = null!;

    /// <summary>Successor task UID.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [OfficeProjectTaskTransformation]
    public OfficeProjectTaskSelector Successor { get; set; } = null!;

    /// <summary>Dependency relationship.</summary>
    [Parameter]
    public ProjectDependencyType Type { get; set; } = ProjectDependencyType.FinishToStart;

    /// <summary>Working-day delay; negative values represent lead time.</summary>
    [Parameter]
    public decimal LagDays { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        Document ??= OfficeDocumentContext.Find<ProjectDocument>() ?? throw new PSInvalidOperationException("Provide -Document or use a Project content block.");
        var dependency = Document.Dependencies.Add(Predecessor.Resolve(Document), Successor.Resolve(Document));
        dependency.Type = Type;
        dependency.Lag = ProjectDuration.WorkingDays(LagDays);
        if (PassThru.IsPresent) WriteObject(dependency);
    }
}

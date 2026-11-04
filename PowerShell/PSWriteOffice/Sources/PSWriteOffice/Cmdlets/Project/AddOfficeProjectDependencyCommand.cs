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
///   <code>$project | Add-OfficeProjectDependency -PredecessorUid 1 -SuccessorUid 2</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeProjectDependency", DefaultParameterSetName = "Document")]
[OutputType(typeof(ProjectDependency))]
public sealed class AddOfficeProjectDependencyCommand : PSCmdlet {
    /// <summary>Project to inspect or change.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;

    /// <summary>Predecessor task UID.</summary>
    [Parameter(Mandatory = true)]
    public int PredecessorUid { get; set; }

    /// <summary>Successor task UID.</summary>
    [Parameter(Mandatory = true)]
    public int SuccessorUid { get; set; }

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
        var dependency = Document.Dependencies.Add(Document.Tasks.GetByUid(PredecessorUid), Document.Tasks.GetByUid(SuccessorUid));
        dependency.Type = Type;
        dependency.Lag = ProjectDuration.WorkingDays(LagDays);
        if (PassThru.IsPresent) WriteObject(dependency);
    }
}

using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Returns project tasks as pipeline objects.</summary>
/// <example>
///   <summary>Returns project tasks as pipeline objects.</summary>
///   <code>$project | Get-OfficeProjectTask | Select-Object Uid, Name, Start, Finish</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeProjectTask", DefaultParameterSetName = "Document")]
[OutputType(typeof(ProjectTask))]
public sealed class GetOfficeProjectTaskCommand : PSCmdlet {
    /// <summary>Project to inspect or change.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;

    /// <summary>Optional stable task UIDs.</summary>
    [Parameter]
    public int[]? Uid { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var tasks = Document.AllTasks;
        if (Uid != null) tasks = tasks.Where(task => Uid.Contains(task.Uid));
        WriteObject(tasks, enumerateCollection: true);
    }
}

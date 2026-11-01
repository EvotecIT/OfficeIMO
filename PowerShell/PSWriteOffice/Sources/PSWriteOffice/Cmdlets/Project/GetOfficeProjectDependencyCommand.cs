using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Returns project dependencies.</summary>
/// <example>
///   <summary>Returns project dependencies.</summary>
///   <code>$project | Get-OfficeProjectDependency</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeProjectDependency", DefaultParameterSetName = "Document")]
[OutputType(typeof(ProjectDependency))]
public sealed class GetOfficeProjectDependencyCommand : PSCmdlet {
    /// <summary>Project document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;
    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var item in Document.Dependencies) {
            WriteObject(item);
        }
    }
}

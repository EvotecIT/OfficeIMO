using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Closes a project without implicitly saving it.</summary>
/// <example>
///   <summary>Closes a project without implicitly saving it.</summary>
///   <code>$project | Close-OfficeProject</code>
/// </example>
[Cmdlet(VerbsCommon.Close, "OfficeProject")]
public sealed class CloseOfficeProjectCommand : PSCmdlet {
    /// <summary>Project to inspect or change.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        Document.Dispose();
    }
}

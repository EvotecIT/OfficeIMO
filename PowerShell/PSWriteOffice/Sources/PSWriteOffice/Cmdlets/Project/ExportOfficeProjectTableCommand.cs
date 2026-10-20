using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Returns mapped project tables and explicit projection-loss notices.</summary>
/// <example>
///   <summary>Returns mapped project tables and explicit projection-loss notices.</summary>
///   <code>$tables = $project | Export-OfficeProjectTable -AllowLossyProjection</code>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeProjectTable")]
[OutputType(typeof(ProjectDataExportResult))]
public sealed class ExportOfficeProjectTableCommand : PSCmdlet {
    /// <summary>Project document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;
    /// <summary>Explicitly accept omitted settings, dependencies and advanced/native data.</summary>
    [Parameter]
    public SwitchParameter AllowLossyProjection { get; set; }
    /// <summary>Maximum total rows.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaxRows { get; set; } = 100000;
    /// <summary>Maximum total cells.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaxCells { get; set; } = 2000000;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var result = Document.ExportTables(AllowLossyProjection.IsPresent, MaxRows, MaxCells);
        foreach (var notice in result.Notices) WriteWarning(notice);
        WriteObject(result);
    }
}

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Imports explicitly mapped tables into a new project without changing existing documents.</summary>
/// <example>
///   <summary>Imports explicitly mapped tables into a new project without changing existing documents.</summary>
///   <code>$taskTable | Import-OfficeProjectTable -Name Delivery -ReportVariable importReport</code>
/// </example>
[Cmdlet(VerbsData.Import, "OfficeProjectTable")]
[OutputType(typeof(ProjectDocument))]
public sealed class ImportOfficeProjectTableCommand : PSCmdlet {
    /// <summary>At most one table for each entity kind.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectMappedTable[] Table { get; set; } = Array.Empty<ProjectMappedTable>();
    /// <summary>Project name.</summary>
    [Parameter]
    public string? Name { get; set; }
    /// <summary>Explicit local start date.</summary>
    [Parameter]
    public DateTime? StartDate { get; set; }
    /// <summary>Imported project calendar UID.</summary>
    [Parameter]
    public int? CalendarUid { get; set; }
    /// <summary>Explicitly permit ignored source columns and return notices.</summary>
    [Parameter]
    public SwitchParameter AllowUnmappedColumns { get; set; }
    /// <summary>Maximum rows across all tables.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaxRows { get; set; } = 100000;
    /// <summary>Variable receiving the import result and notices.</summary>
    [Parameter]
    public string? ReportVariable { get; set; }

    private readonly List<ProjectMappedTable> _tables = new();
    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (_tables.Count + Table.Length > 4) throw new PSArgumentException("Supply at most four mapped tables.");
        _tables.AddRange(Table);
    }
    /// <inheritdoc />
    protected override void EndProcessing() {
        var result = ProjectDocument.ImportTables(_tables, new ProjectDataImportOptions {
            Name = Name, Start = StartDate, CalendarUid = CalendarUid,
            AllowUnmappedColumns = AllowUnmappedColumns.IsPresent, MaxRows = MaxRows
        });
        if (!string.IsNullOrWhiteSpace(ReportVariable)) SessionState.PSVariable.Set(ReportVariable, result);
        foreach (var notice in result.Notices) WriteWarning(notice);
        WriteObject(result.Document);
    }
}

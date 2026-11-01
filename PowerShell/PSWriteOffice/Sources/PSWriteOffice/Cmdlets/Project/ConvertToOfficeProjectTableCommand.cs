using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Maps PowerShell rows to a bounded project entity table.</summary>
/// <example>
///   <summary>Maps PowerShell rows to a bounded project entity table.</summary>
///   <code>Import-Csv ./tasks.csv | ConvertTo-OfficeProjectTable -Kind Tasks -ColumnMap @{ Uid = 'ID'; Name = 'Task' }</code>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeProjectTable", DefaultParameterSetName = "InputObject")]
[OutputType(typeof(ProjectMappedTable))]
public sealed class ConvertToOfficeProjectTableCommand : PSCmdlet {
    /// <summary>Rows from Import-Csv, dictionaries or PowerShell objects.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public PSObject[] InputObject { get; set; } = Array.Empty<PSObject>();
    /// <summary>Entity collection represented by the rows.</summary>
    [Parameter(Mandatory = true)]
    public ProjectDataKind Kind { get; set; }
    /// <summary>Project field name to source column header. Field names are discovered through ProjectDataField.</summary>
    [Parameter(Mandatory = true)]
    public System.Collections.IDictionary ColumnMap { get; set; } = null!;
    /// <summary>Maximum accumulated rows.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaxRows { get; set; } = 100000;

    private readonly List<object?> _rows = new();
    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var row in InputObject) {
            if (_rows.Count >= MaxRows) throw new PSArgumentException("Input exceeds MaxRows.");
            _rows.Add(row);
        }
    }
    /// <inheritdoc />
    protected override void EndProcessing() {
        using var table = PSWriteOffice.Services.Excel.ExcelTabularInputService.ToDataTable(_rows);
        var headers = table.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName).ToArray();
        var rows = table.Rows.Cast<System.Data.DataRow>().Select(row => headers.Select((_, index) => {
            var value = row[index];
            return value == DBNull.Value ? null : value is DateTime date ? date.ToString("s", System.Globalization.CultureInfo.InvariantCulture)
                : Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
        }).ToArray());
        var columns = new List<ProjectDataColumn>();
        foreach (System.Collections.DictionaryEntry item in ColumnMap) {
            if (!Enum.TryParse<ProjectDataField>(Convert.ToString(item.Key), true, out var field) || !Enum.IsDefined(typeof(ProjectDataField), field))
                throw new PSArgumentException("Unknown project field: " + item.Key);
            columns.Add(new ProjectDataColumn(field, Convert.ToString(item.Value) ?? string.Empty));
        }
        WriteObject(new ProjectMappedTable(Kind, new ProjectDataTable(headers, rows, MaxRows), columns));
    }
}

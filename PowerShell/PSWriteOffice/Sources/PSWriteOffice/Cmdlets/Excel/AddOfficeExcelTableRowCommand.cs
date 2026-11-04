using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Table;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Appends rows to an existing Excel table.</summary>
/// <example>
/// <summary>Append rows to a table</summary>
/// <code>$servers = @(
///     [pscustomobject]@{ Name = 'NW-APP-07'; Role = 'Application'; Site = 'Warsaw' }
///     [pscustomobject]@{ Name = 'NW-SQL-03'; Role = 'Database'; Site = 'Frankfurt' }
/// )
/// Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
///     Add-OfficeExcelTableRow -Sheet Servers -Table Servers -InputObject $servers
/// }</code>
/// <para>Adds two new servers to the Servers table; the table range grows to include them.</para>
/// </example>
/// <example>
/// <summary>Log each deployment as it finishes</summary>
/// <code>$book = Open-OfficeExcel -Path .\Reports\Deployments.xlsx
/// foreach ($target in 'NW-APP-01', 'NW-APP-02', 'NW-APP-03') {
///     $result = Invoke-Command -ComputerName $target { &amp; C:\Deploy\Install.ps1 -Release '2026.10.2' }
///     [pscustomobject]@{ Server = $target; Release = '2026.10.2'; Finished = Get-Date; Status = $result.Status } |
///         Add-OfficeExcelTableRow -Document $book -Sheet Log -Table Deployments
/// }
/// Save-OfficeExcel -Document $book -PassThru | Close-OfficeExcel</code>
/// <para>Appends one row per completed deployment to the Deployments table of a workbook kept open for the whole run.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelTableRow", DefaultParameterSetName = "Default")]
public sealed class AddOfficeExcelTableRowCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Native Excel table used to resolve the source range or header.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true)]
    public object Table { get; set; } = null!;

    /// <summary>Objects, dictionaries or native tabular data supplied directly or through the pipeline.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 1, ValueFromPipeline = true)]
    public object InputObject { get; set; } = null!;

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        object? target = OfficeArgumentValue.Unwrap(Table);
        ExcelTable table;
        if (target is ExcelTable native) {
            if (Document != null && !ReferenceEquals(native.Sheet.Document, Document)) throw new PSArgumentException("The table does not belong to Document.");
            table = native;
        } else {
            var document = ExcelCommandTargets.Document(Document);
            string name = Convert.ToString(target)!;
            var sheets = Sheet == null ? document.Sheets : new[] { Sheet.Resolve(document.Sheets, s => s.Name) };
            var matches = sheets.Where(s => s.GetTableRange(name) != null).ToArray();
            if (matches.Length != 1) throw new PSArgumentException("Table must identify one existing table; use Sheet to narrow it.");
            table = matches[0].Table(name);
        }
        var rows = new List<object?>();
        TableInputCollector.AddInput(rows, InputObject, preserveTabularInput: true);
        table.AppendDataTable(ExcelTabularInputService.ToDataTable(rows, table.NameOrRange));
        if (PassThru) WriteObject(table);
    }

}

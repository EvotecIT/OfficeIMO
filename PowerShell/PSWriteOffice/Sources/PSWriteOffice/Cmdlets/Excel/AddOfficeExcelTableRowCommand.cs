using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Appends one or more data rows to an existing Excel table.</summary>
/// <para>
/// Use this command when a workbook already contains a named table and the script should extend that
/// table without recreating the worksheet through the Excel DSL. The command accepts a workbook path,
/// an open <see cref="ExcelDocument"/>, or an existing <see cref="ExcelTable"/> object. Objects,
/// dictionaries, <c>DataTable</c>, <c>DataView</c>, <c>IDataReader</c>, and <c>DataRow</c> input are
/// normalized through the same table input pipeline used by <c>Add-OfficeExcelTable</c>.
/// </para>
/// <para>
/// When a path is supplied, PSWriteOffice opens the workbook, appends the rows, saves the workbook, and
/// releases the document. When an open workbook or table is supplied, the caller controls the lifetime
/// and should close or save the workbook after all edits are complete.
/// </para>
/// <example>
///   <summary>Append a row to a named table in an open workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeExcel -Path .\Report.xlsx
/// $doc | Add-OfficeExcelTableRow -Sheet Data -TableName Sales -InputObject ([pscustomobject]@{ Region='APAC'; Revenue=300 })
/// $doc | Close-OfficeExcel -Save</code>
///   <para>Uses the existing OfficeIMO Excel table append API and keeps the workbook open for further changes.</para>
/// </example>
/// <example>
///   <summary>Append several service-readiness rows to an existing table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = @(
///     [pscustomobject]@{ Service='Identity'; Status='Ready'; Owner='IAM' }
///     [pscustomobject]@{ Service='Network'; Status='Investigating'; Owner='Platform' }
/// )
/// Add-OfficeExcelTableRow -Path .\Readiness.xlsx -Sheet Readiness -TableName ServiceReadiness -InputObject $rows</code>
///   <para>Opens the workbook from disk, appends both objects to the named table, and saves the file.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelTableRow", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[OutputType(typeof(ExcelTable))]
public sealed class AddOfficeExcelTableRowCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetTable = "Table";

    private readonly List<object?> _items = new();

    /// <summary>Workbook path to open, update, save, and close.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open workbook to update. The caller remains responsible for saving and closing it.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Existing OfficeIMO Excel table wrapper to append to when the table has already been resolved.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetTable)]
    public ExcelTable? Table { get; set; }

    /// <summary>Existing table name or display name, for example the name returned by <c>Get-OfficeExcelTable</c>.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPath)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocument)]
    [Alias("Name")]
    public string TableName { get; set; } = string.Empty;

    /// <summary>Worksheet name that owns the table. Use this when table names might repeat across sheets.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Zero-based worksheet index that owns the table.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>Rows to append. Accepts objects, dictionaries, DataTables, DataViews, IDataReaders, and DataRows.</summary>
    [Parameter(Mandatory = true, Position = 1, ValueFromPipeline = true)]
    [Alias("Data", "Values")]
    public object? InputObject { get; set; }

    /// <summary>
    /// Emit the updated table wrapper for open document or table inputs. Path-owned workbooks are saved and closed by this command,
    /// so they do not emit a live table wrapper.
    /// </summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (string.Equals(ParameterSetName, ParameterSetPath, StringComparison.OrdinalIgnoreCase)) {
            TableInputCollector.AddInput(_items, InputObject, preserveTabularInput: true);
            return;
        }

        AppendRows(CreateDataTable(InputObject), saveOwnedWorkbook: false);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (!string.Equals(ParameterSetName, ParameterSetPath, StringComparison.OrdinalIgnoreCase)) {
            return;
        }

        AppendRows(ExcelTabularInputService.ToDataTable(_items, TableName), saveOwnedWorkbook: true);
    }

    private void AppendRows(System.Data.DataTable data, bool saveOwnedWorkbook) {
        ExcelTable table;

        if (ParameterSetName == ParameterSetTable) {
            table = Table ?? throw new PSArgumentException("Provide an Excel table.", nameof(Table));
            if (!ExcelShouldProcessService.ShouldProcessTarget(this, "Excel table", "Append Excel table rows")) {
                return;
            }

            table.AppendDataTable(data);
        } else {
            using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
            if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Append Excel table rows")) {
                return;
            }

            table = ResolveTable(workbook.Document);
            table.AppendDataTable(data);
            if (saveOwnedWorkbook) {
                workbook.SaveIfOwned();
            }
        }

        if (PassThru.IsPresent) {
            if (string.Equals(ParameterSetName, ParameterSetPath, StringComparison.OrdinalIgnoreCase)) {
                WriteWarning("Path-owned workbooks are saved and closed by Add-OfficeExcelTableRow; no live ExcelTable is emitted. Open the workbook with Get-OfficeExcel when chaining table edits.");
                return;
            }

            WriteObject(table);
        }
    }

    private System.Data.DataTable CreateDataTable(object? value) {
        var items = new List<object?>();
        TableInputCollector.AddInput(items, value, preserveTabularInput: true);
        return ExcelTabularInputService.ToDataTable(items, TableName);
    }

    private ExcelTable ResolveTable(ExcelDocument document) {
        if (!string.IsNullOrWhiteSpace(Sheet) || SheetIndex.HasValue) {
            var sheet = ExcelWorkbookCommandService.ResolveSheet(this, document, ParameterSetName, Sheet, SheetIndex);
            return sheet.Table(TableName);
        }

        var matches = document.Sheets
            .Where(sheet => sheet.GetTableRange(TableName) != null)
            .ToArray();

        if (matches.Length == 0) {
            throw new PSArgumentException($"Table '{TableName}' was not found in the workbook.", nameof(TableName));
        }

        if (matches.Length > 1) {
            throw new PSArgumentException(
                $"Table '{TableName}' exists on multiple worksheets. Specify -Sheet or -SheetIndex to select one.",
                nameof(TableName));
        }

        return matches[0].Table(TableName);
    }
}
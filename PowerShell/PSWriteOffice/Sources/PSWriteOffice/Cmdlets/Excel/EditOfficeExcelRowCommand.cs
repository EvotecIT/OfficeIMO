using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Runs a script block against editable worksheet rows.</summary>
/// <example>
///   <summary>Edit rows by header name.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Edit-OfficeExcelRow -Path .\Report.xlsx -Sheet Data -ScriptBlock { param($row) if ($row.Get[string]('Status') -eq 'Draft') { $row.Set('Status', 'Ready') } }</code>
///   <para>Loads editable row handles, lets the script update cells, and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsData.Edit, "OfficeExcelRow", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("Edit-ExcelRow", "ExcelRowEdit")]
[OutputType(typeof(ExcelPowerShellRowEdit))]
public sealed class EditOfficeExcelRowCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>A1 range to expose as editable rows. Defaults to the worksheet used range.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Script block to run once per editable row. The row is passed as the first argument.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public ScriptBlock ScriptBlock { get; set; } = null!;

    /// <summary>Prefer decimals instead of doubles for numeric values.</summary>
    [Parameter]
    public SwitchParameter NumericAsDecimal { get; set; }

    /// <summary>Emit each editable row after the script block runs.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {
            return;
        }

        var document = workbook.Document;
        var sheet = ExcelWorkbookCommandService.ResolveSheet(this, document, ParameterSetName, Sheet, SheetIndex);
        var options = ExcelReadOutputService.CreateOptions(NumericAsDecimal.IsPresent);
        var range = string.IsNullOrWhiteSpace(Range) ? sheet.UsedRangeA1 : Range!;
        var (firstRow, firstColumn, _, _) = A1.ParseRange(range);
        options.SheetName = sheet.Name;
        options.A1Range = range;
        options.HasHeaderRow = true;

        using var reader = document.CreateDataReader(options);
        var headers = new string[reader.FieldCount];
        for (var columnIndex = 0; columnIndex < headers.Length; columnIndex++) {
            headers[columnIndex] = reader.GetName(columnIndex);
        }

        var dataRowIndex = 0;
        while (reader.Read()) {
            var values = new object[reader.FieldCount];
            reader.GetValues(values);
            var row = new ExcelPowerShellRowEdit(
                sheet,
                firstRow + 1 + dataRowIndex,
                firstColumn,
                headers,
                values,
                options.Culture);
            ScriptBlock.Invoke(row);
            if (PassThru.IsPresent) {
                WriteObject(row);
            }

            dataRowIndex++;
        }

        workbook.SaveIfOwned();
    }
}
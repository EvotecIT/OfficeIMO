using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Runs a script block against each editable row of a range and writes the changes back.</summary>
/// <example>
/// <summary>Update rows in place</summary>
/// <code>Edit-OfficeExcelRow -Path .\Reports\Releases.xlsx -Sheet Releases -ScriptBlock {
///     param($row)
///     if ($row['Status'].Value -eq 'Draft') { $row.Set('Status', 'Ready') }
/// }</code>
/// <para>Marks every draft row on the Releases sheet as ready and saves the workbook.</para>
/// </example>
/// <example>
/// <summary>Recalculate renewal dates in an open workbook</summary>
/// <code>$book = Open-OfficeExcel -Path .\Contracts\Licences.xlsx
/// Edit-OfficeExcelRow -Document $book -Sheet Licences -Range A1:F400 -ScriptBlock {
///     param($row)
///     if ($row['State'].Value -eq 'Active') {
///         $row.Set('RenewalDate', $row.Get[datetime]('StartDate').AddYears(1))
///     }
/// }
/// Save-OfficeExcel -Document $book -PassThru | Close-OfficeExcel</code>
/// <para>Sets the renewal date of each active licence to one year after its start date in a live workbook, which the script then saves.</para>
/// </example>
/// <example>
/// <summary>Edit rows while building a workbook</summary>
/// <code>$tickets = Import-OfficeCsv -Path .\Data\Tickets.csv
/// New-OfficeExcel -Path .\Reports\Tickets.xlsx {
///     ExcelSheet 'Tickets' {
///         ExcelTable $tickets -Name Tickets
///     }
///     Edit-OfficeExcelRow -Sheet Tickets -ScriptBlock {
///         param($row)
///         if (-not $row['Owner'].Value) { $row.Set('Owner', 'Unassigned') }
///     }
/// }</code>
/// <para>Writes imported rows and then fixes up their status inside the same block, without a -Document variable.</para>
/// </example>
[Cmdlet(VerbsData.Edit, "OfficeExcelRow", SupportsShouldProcess = true, DefaultParameterSetName = "Document")]
public sealed class EditOfficeExcelRowCommand : PSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter()]
    public string? Range { get; set; }

    /// <summary>Script block to run once per editable row. The row is passed as the first argument.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public ScriptBlock ScriptBlock { get; set; } = null!;

    /// <summary>Prefer decimals instead of doubles for numeric values.</summary>
    [Parameter()]
    public SwitchParameter NumericAsDecimal { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Literal workbook paths, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var workbook in ExcelCommandTargets.Workbooks(this, ParameterSetName, Path, LiteralPath, Document, readOnly: false)) {
            if (workbook.OwnsDocument && !ShouldProcess(workbook.Document.FilePath, "Edit Excel rows")) continue;
            var document = workbook.Document;
            var sheet = ExcelSheetResolver.ResolveTarget(this, document, Sheet);
        var options = ExcelReadOutputService.CreateOptions(NumericAsDecimal.IsPresent);
        var range = string.IsNullOrWhiteSpace(Range) ? sheet.UsedRange.Address : Range!;
        var (firstRow, firstColumn, _, _) = ExcelA1.ParseRange(range);
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

}

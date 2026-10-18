using System;
using System.Globalization;
using System.Management.Automation;
using System.Text.RegularExpressions;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Finds text in worksheet values.</summary>
/// <example>
///   <summary>Find status values and export their addresses.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$matches = Find-OfficeExcel -Path .\Report.xlsx -Text Ready -Sheet Summary
/// $matches |
///     Select-Object -Property Sheet, Address, Value |
///     Export-Csv -Path .\ReadyCells.csv -NoTypeInformation</code>
///   <para>Returns matching cells with sheet, address, row, column, and value metadata for review or proof.</para>
/// </example>
[Cmdlet(VerbsCommon.Find, "OfficeExcel", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(PSObject))]
public sealed class FindOfficeExcelCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to inspect.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to inspect outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Text or pattern to find.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Worksheet name. Defaults to all sheets for path/document use and current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>A1 range to search. Defaults to each selected worksheet's used range.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Use case-sensitive matching.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Treat -Text as a regular expression.</summary>
    [Parameter]
    public SwitchParameter Regex { get; set; }

    /// <summary>Require an exact cell text match instead of substring matching.</summary>
    [Parameter]
    public SwitchParameter Exact { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var document = workbook.Document;
        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, document, ParameterSetName, Sheet, SheetIndex)) {
            var range = string.IsNullOrWhiteSpace(Range) ? sheet.UsedRangeA1 : Range!;
            foreach (var cell in sheet.EnumerateRange(range)) {
                var cellText = Convert.ToString(cell.Value, CultureInfo.InvariantCulture) ?? string.Empty;
                if (IsMatch(cellText)) {
                    WriteObject(CreateRecord(sheet.Name, cell.Row, cell.Column, cell.Value));
                }
            }
        }
    }

    private bool IsMatch(string value) {
        var comparison = CaseSensitive.IsPresent ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        if (Regex.IsPresent) {
            var options = CaseSensitive.IsPresent ? RegexOptions.None : RegexOptions.IgnoreCase;
            return System.Text.RegularExpressions.Regex.IsMatch(value, Text, options);
        }

        return Exact.IsPresent
            ? string.Equals(value, Text, comparison)
            : value.IndexOf(Text, comparison) >= 0;
    }

    private static PSObject CreateRecord(string sheetName, int row, int column, object? value) {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Sheet", sheetName));
        record.Properties.Add(new PSNoteProperty("Address", A1.CellReference(row, column)));
        record.Properties.Add(new PSNoteProperty("Row", row));
        record.Properties.Add(new PSNoteProperty("Column", column));
        record.Properties.Add(new PSNoteProperty("Value", value));
        return record;
    }
}
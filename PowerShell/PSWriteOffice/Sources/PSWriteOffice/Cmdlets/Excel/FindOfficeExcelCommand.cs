using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using System.Text.RegularExpressions;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Finds cells whose values match a pattern in a workbook, sheet or range.</summary>
/// <example>
/// <summary>Find cells containing a value</summary>
/// <code>Find-OfficeExcel -Path .\Reports\Inventory.xlsx -Pattern 'WEB01' -SimpleMatch</code>
/// <para>Lists every cell in the workbook that mentions the server WEB01.</para>
/// </example>
/// <example>
/// <summary>Find retired certificates that are still listed</summary>
/// <code>$old = Import-OfficeCsv -Path .\Security\Retired-Certificates.csv
/// foreach ($certificate in $old) {
///     Find-OfficeExcel -Path .\Security\Certificates.xlsx -Sheet Inventory -Pattern $certificate.Thumbprint -SimpleMatch -Exact
/// }</code>
/// <para>Looks up each retired certificate thumbprint on the inventory sheet and shows where it still appears.</para>
/// </example>
[Cmdlet(VerbsCommon.Find, "OfficeExcel", DefaultParameterSetName = "Document")]
public sealed class FindOfficeExcelCommand : PSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Date or DateTime number format pattern.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Pattern { get; set; } = null!;

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    [Alias("WorksheetName")]
    public OfficeSheetSelector[]? Sheet { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter()]
    public string? Range { get; set; }

    /// <summary>Uses literal text matching instead of regular expressions.</summary>
    [Parameter()]
    public SwitchParameter SimpleMatch { get; set; }

    /// <summary>Matches names and text using ordinal case sensitivity. Omission ignores case.</summary>
    [Parameter()]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Require an exact cell text match instead of substring matching.</summary>
    [Parameter()]
    public SwitchParameter Exact { get; set; }

    /// <summary>Literal workbook paths, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var scope in ExcelCommandTargets.Workbooks(this, ParameterSetName, Path, LiteralPath, Document, readOnly: true)) {
            foreach (var sheet in ExcelSheetResolver.ResolveMany(scope.Document, Sheet, Document == null && !scope.OwnsDocument ? ExcelDslContext.Current?.CurrentSheet : null))
                foreach (var cell in sheet.EnumerateRange(Range ?? sheet.UsedRange.Address)) {
                    if (IsMatch(Convert.ToString(cell.Value, CultureInfo.InvariantCulture) ?? string.Empty)) WriteObject(CreateRecord(sheet.Name, cell.Row, cell.Column, cell.Value));
                }
        }
    }
    private bool IsMatch(string value) {
        var comparison = CaseSensitive.IsPresent ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        if (!SimpleMatch.IsPresent) {
            var options = CaseSensitive.IsPresent ? RegexOptions.None : RegexOptions.IgnoreCase;
            return System.Text.RegularExpressions.Regex.IsMatch(value, Pattern, options);
        }

        return Exact.IsPresent
            ? string.Equals(value, Pattern, comparison)
            : value.IndexOf(Pattern, comparison) >= 0;
    }

    private static PSObject CreateRecord(string sheetName, int row, int column, object? value) {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Sheet", sheetName));
        record.Properties.Add(new PSNoteProperty("Address", ExcelA1.CellReference(row, column)));
        record.Properties.Add(new PSNoteProperty("Row", row));
        record.Properties.Add(new PSNoteProperty("Column", column));
        record.Properties.Add(new PSNoteProperty("Value", value));
        return record;
    }

}

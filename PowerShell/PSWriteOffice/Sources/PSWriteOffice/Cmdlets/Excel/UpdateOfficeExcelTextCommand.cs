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

/// <summary>Replaces text in cell values across a workbook, sheet or range.</summary>
/// <example>
/// <summary>Replace text across a workbook</summary>
/// <code>Update-OfficeExcelText -Path .\Templates\Price-List.xlsx -Pattern 'Fabrikam Ltd' -Replacement 'Northwind Traders' -SimpleMatch</code>
/// <para>Replaces the old company name in every cell of the workbook and saves it.</para>
/// </example>
/// <example>
/// <summary>Rename hosts after a migration</summary>
/// <code>$changed = Update-OfficeExcelText -Path .\Reports\Inventory.xlsx -Sheet Servers, Services -Pattern '^SRV-(\d+)$' -Replacement 'NW-APP-$1' -PassThru
/// "Renamed $changed cells"</code>
/// <para>Uses a regular expression to rename the old SRV hosts to the new naming scheme on the inventory sheets and reports how many cells changed.</para>
/// </example>
/// <example>
/// <summary>Replace placeholders inside a block</summary>
/// <code>Open-OfficeExcel -Path .\Templates\Quote.xlsx {
///     Update-OfficeExcelText -Pattern '{{Customer}}' -Replacement 'Northwind Traders' -SimpleMatch
///     Update-OfficeExcelText -Pattern '{{Date}}' -Replacement (Get-Date -Format 'yyyy-MM-dd') -SimpleMatch
/// }</code>
/// <para>Fills a template and replaces its placeholder text in the same Open-OfficeExcel block.</para>
/// </example>
[Cmdlet(VerbsData.Update, "OfficeExcelText", SupportsShouldProcess = true, DefaultParameterSetName = "Document")]
public sealed class UpdateOfficeExcelTextCommand : PSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Date or DateTime number format pattern.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Pattern { get; set; } = null!;

    /// <summary>Replacement text or regular-expression replacement pattern.</summary>
    [Parameter(Mandatory = true, Position = 2)]
    public string Replacement { get; set; } = null!;

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

    /// <summary>Open the saved file in its registered application after a local block completes.</summary>
    [Parameter()]
    public SwitchParameter Show { get; set; }

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
        if (Pattern.Length == 0) throw new PSArgumentException("Pattern cannot be empty.");
        foreach (var scope in ExcelCommandTargets.Workbooks(this, ParameterSetName, Path, LiteralPath, Document, readOnly: false)) {
            if (scope.OwnsDocument && !ShouldProcess(scope.Document.FilePath, "Replace Excel text")) continue;
            int count = 0;
            foreach (var sheet in ExcelSheetResolver.ResolveMany(scope.Document, Sheet, Document == null && !scope.OwnsDocument ? ExcelDslContext.Current?.CurrentSheet : null)) count += ReplaceInSheet(sheet);
            scope.SaveIfOwned();
            if (Show && scope.OwnsDocument) FileOpenService.Open(scope.Document.FilePath!);
            if (PassThru) WriteObject(count);
        }
    }
    private int ReplaceInSheet(ExcelSheet sheet) {
        var count = 0;
        var range = string.IsNullOrWhiteSpace(Range) ? sheet.UsedRange.Address : Range!;
        foreach (var cell in sheet.EnumerateRange(range)) {
            if (cell.Value is not string text) {
                continue;
            }

            var updated = ReplaceString(text, out var cellReplacements);
            if (cellReplacements == 0) {
                continue;
            }

            sheet.Cell(cell.Row, cell.Column).SetValue(updated);
            count += cellReplacements;
        }

        return count;
    }

    private string ReplaceString(string value, out int replacements) {
        replacements = 0;
        if (!SimpleMatch.IsPresent) {
            var options = CaseSensitive.IsPresent ? RegexOptions.None : RegexOptions.IgnoreCase;
            var count = 0;
            var updated = System.Text.RegularExpressions.Regex.Replace(
                value,
                Pattern,
                match => {
                    count++;
                    return Replacement;
                },
                options);
            replacements = count;
            return updated;
        }

        var comparison = CaseSensitive.IsPresent ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var index = value.IndexOf(Pattern, comparison);
        if (index < 0) {
            return value;
        }

        var result = value;
        while (index >= 0) {
            result = result.Substring(0, index) + Replacement + result.Substring(index + Pattern.Length);
            replacements++;
            index = result.IndexOf(Pattern, index + Replacement.Length, comparison);
        }

        return result;
    }


}

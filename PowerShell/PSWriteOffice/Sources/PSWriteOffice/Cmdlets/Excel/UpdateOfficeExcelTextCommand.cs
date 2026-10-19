using System;
using System.Management.Automation;
using System.Text.RegularExpressions;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Replaces text in worksheet values.</summary>
/// <example>
///   <summary>Replace status text and verify the update count.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$count = Update-OfficeExcelText -Path .\Report.xlsx -Sheet Summary -OldValue Draft -NewValue Ready -PassThru
/// [pscustomobject]@{
///     Path = '.\Report.xlsx'
///     Replacements = $count
/// }</code>
///   <para>Updates matching text cells on a sheet, saves the workbook, and returns the replacement count.</para>
/// </example>
[Cmdlet(VerbsData.Update, "OfficeExcelText", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("Replace-OfficeExcelText")]
[OutputType(typeof(int))]
public sealed class UpdateOfficeExcelTextCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
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

    /// <summary>Text or pattern to replace.</summary>
    [Parameter(Mandatory = true)]
    public string OldValue { get; set; } = string.Empty;

    /// <summary>Replacement text.</summary>
    [Parameter(Mandatory = true)]
    public string NewValue { get; set; } = string.Empty;

    /// <summary>Worksheet name. Defaults to all sheets for path/document use and current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>A1 range to update. Defaults to each selected worksheet's used range.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Use case-sensitive matching.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Treat -OldValue as a regular expression.</summary>
    [Parameter]
    public SwitchParameter Regex { get; set; }

    /// <summary>Open the file after saving when using -Path.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == ParameterSetPath) {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!ShouldProcess(resolvedPath, "Update Excel workbook text")) {
                return;
            }
        }

        var replacements = 0;
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        var document = workbook.Document;
        if (ParameterSetName != ParameterSetPath &&
            !ExcelShouldProcessService.ShouldProcessWorkbook(this, document, null, "Update Excel workbook text")) {
            return;
        }

        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, document, ParameterSetName, Sheet, SheetIndex)) {
            replacements += ReplaceInSheet(sheet);
        }

        workbook.SaveIfOwned();
        var openPath = workbook.OwnsDocument && Open.IsPresent
            ? document.FilePath ?? Path
            : null;

        if (!string.IsNullOrWhiteSpace(openPath)) {
            workbook.Dispose();
            FileOpenService.Open(openPath!);
        }

        WritePassThru(replacements);
    }

    private int ReplaceInSheet(ExcelSheet sheet) {
        var count = 0;
        var range = string.IsNullOrWhiteSpace(Range) ? sheet.UsedRangeA1 : Range!;
        foreach (var cell in sheet.EnumerateRange(range)) {
            if (cell.Value is not string text) {
                continue;
            }

            var updated = ReplaceString(text, out var cellReplacements);
            if (cellReplacements == 0) {
                continue;
            }

            sheet.CellValue(cell.Row, cell.Column, updated);
            count += cellReplacements;
        }

        return count;
    }

    private string ReplaceString(string value, out int replacements) {
        replacements = 0;
        if (Regex.IsPresent) {
            var options = CaseSensitive.IsPresent ? RegexOptions.None : RegexOptions.IgnoreCase;
            var count = 0;
            var updated = System.Text.RegularExpressions.Regex.Replace(
                value,
                OldValue,
                match => {
                    count++;
                    return NewValue;
                },
                options);
            replacements = count;
            return updated;
        }

        var comparison = CaseSensitive.IsPresent ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var index = value.IndexOf(OldValue, comparison);
        if (index < 0) {
            return value;
        }

        var result = value;
        while (index >= 0) {
            result = result.Substring(0, index) + NewValue + result.Substring(index + OldValue.Length);
            replacements++;
            index = result.IndexOf(OldValue, index + NewValue.Length, comparison);
        }

        return result;
    }

}

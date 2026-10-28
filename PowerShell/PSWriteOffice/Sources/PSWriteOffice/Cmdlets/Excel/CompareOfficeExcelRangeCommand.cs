using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Compares two Excel worksheets or ranges and returns cell-level differences.</summary>
/// <example>
///   <summary>Compare current and expected ranges before publishing.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$differences = Compare-OfficeExcelRange -Path .\Report.xlsx -LeftSheet Current -RightSheet Expected -LeftRange A1:D20 -RightRange A1:D20 -TrimStrings -IgnoreCase
/// if ($differences) {
///     $differences | Export-Csv -Path .\RangeDifferences.csv -NoTypeInformation
/// }</code>
///   <para>Compares two ranges and exports cell-level differences when the workbook does not match the expected sheet.</para>
/// </example>
[Cmdlet(VerbsData.Compare, "OfficeExcelRange", DefaultParameterSetName = ParameterSetPath)]
[Alias("Compare-OfficeExcelSheet", "ExcelCompare")]
[OutputType(typeof(ExcelRangeDifference))]
public sealed class CompareOfficeExcelRangeCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Left workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath", "LeftPath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Left workbook object.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Optional right workbook path. Defaults to the left workbook.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? RightPath { get; set; }

    /// <summary>Optional right workbook object. Defaults to the left workbook.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public ExcelDocument? RightDocument { get; set; }

    /// <summary>Left worksheet name. Defaults to the current sheet in a sheet block, otherwise the first sheet.</summary>
    [Parameter]
    public string? LeftSheet { get; set; }

    /// <summary>Left worksheet index.</summary>
    [Parameter]
    public int? LeftSheetIndex { get; set; }

    /// <summary>Right worksheet name. Defaults to the left sheet.</summary>
    [Parameter]
    public string? RightSheet { get; set; }

    /// <summary>Right worksheet index.</summary>
    [Parameter]
    public int? RightSheetIndex { get; set; }

    /// <summary>Left A1 range. Defaults to the left worksheet used range.</summary>
    [Parameter]
    public string? LeftRange { get; set; }

    /// <summary>Right A1 range. Defaults to the right worksheet used range.</summary>
    [Parameter]
    public string? RightRange { get; set; }

    /// <summary>Compare strings after trimming whitespace.</summary>
    [Parameter]
    public SwitchParameter TrimStrings { get; set; }

    /// <summary>Compare strings case-insensitively.</summary>
    [Parameter]
    public SwitchParameter IgnoreCase { get; set; }

    /// <summary>Treat null and empty string values as different.</summary>
    [Parameter]
    public SwitchParameter StrictNullEmpty { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var leftWorkbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var leftDocument = leftWorkbook.Document;
        using var rightWorkbook = ResolveRightWorkbook(leftDocument);
        var rightDocument = rightWorkbook.Document;
        var leftSheet = ExcelWorkbookCommandService.ResolveSheet(this, leftDocument, ParameterSetName, LeftSheet, LeftSheetIndex);
        var rightSheet = ResolveRightSheet(rightDocument, leftSheet.Name);

        var options = new ExcelRangeCompareOptions {
            TrimStrings = TrimStrings.IsPresent,
            IgnoreCase = IgnoreCase.IsPresent,
            TreatNullAndEmptyStringAsEqual = !StrictNullEmpty.IsPresent
        };

        var differences = leftDocument.CompareRanges(
            leftSheet,
            string.IsNullOrWhiteSpace(LeftRange) ? leftSheet.UsedRangeA1 : LeftRange!,
            rightSheet,
            string.IsNullOrWhiteSpace(RightRange) ? rightSheet.UsedRangeA1 : RightRange!,
            options);

        WriteObject(differences, enumerateCollection: true);
    }

    private ExcelWorkbookCommandScope ResolveRightWorkbook(ExcelDocument leftDocument) {
        if (ParameterSetName == ParameterSetPath && !string.IsNullOrWhiteSpace(RightPath)) {
            return ExcelWorkbookCommandService.OpenWorkbook(this, RightPath!, readOnly: true);
        }

        return new ExcelWorkbookCommandScope(RightDocument ?? leftDocument, ownsDocument: false);
    }

    private ExcelSheet ResolveRightSheet(ExcelDocument document, string leftSheetName) {
        if (!string.IsNullOrWhiteSpace(RightSheet) || RightSheetIndex.HasValue) {
            return ExcelSheetResolver.Resolve(document, RightSheet, RightSheetIndex);
        }

        return document[leftSheetName];
    }
}
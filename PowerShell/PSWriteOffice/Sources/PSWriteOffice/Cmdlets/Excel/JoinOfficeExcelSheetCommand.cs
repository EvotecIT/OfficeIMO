using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Appends or merges rows from one worksheet into another.</summary>
/// <example>
///   <summary>Append source rows below a combined sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$result = Join-OfficeExcelSheet -Path .\Report.xlsx -TargetSheet Combined -SourceSheet Data -MatchColumnsByHeader -BlankRowsBefore 1
/// $result |
///     Select-Object -Property SourceSheet, TargetSheet, RowsCopied, ColumnsCopied</code>
///   <para>Copies rows from Data into Combined, aligns columns by header, and returns the merge result.</para>
/// </example>
[Cmdlet(VerbsCommon.Join, "OfficeExcelSheet", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("Merge-OfficeExcelSheet", "ExcelSheetJoin", "ExcelSheetMerge")]
[OutputType(typeof(ExcelWorksheetMergeResult))]
public sealed class JoinOfficeExcelSheetCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Target workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Target workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Target worksheet name. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    public string? TargetSheet { get; set; }

    /// <summary>Target worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? TargetSheetIndex { get; set; }

    /// <summary>Source worksheet name.</summary>
    [Parameter(Mandatory = true)]
    public string SourceSheet { get; set; } = string.Empty;

    /// <summary>Optional source workbook object for cross-workbook joins.</summary>
    [Parameter]
    public ExcelDocument? SourceDocument { get; set; }

    /// <summary>Optional source workbook path for cross-workbook joins.</summary>
    [Parameter]
    public string? SourcePath { get; set; }

    /// <summary>Source A1 range to copy. Defaults to the source used range.</summary>
    [Parameter]
    public string? SourceRange { get; set; }

    /// <summary>1-based target start row. Defaults to appending after the target used range.</summary>
    [Parameter]
    public int? TargetStartRow { get; set; }

    /// <summary>1-based target start column. Defaults to the source range start column.</summary>
    [Parameter]
    public int? TargetStartColumn { get; set; }

    /// <summary>Treat the first source row as data instead of a header row.</summary>
    [Parameter]
    public SwitchParameter NoSourceHeader { get; set; }

    /// <summary>Include the source header row in copied rows.</summary>
    [Parameter]
    public SwitchParameter IncludeSourceHeader { get; set; }

    /// <summary>Match source columns to target columns by header text.</summary>
    [Parameter]
    public SwitchParameter MatchColumnsByHeader { get; set; }

    /// <summary>1-based target header row when matching columns by header.</summary>
    [Parameter]
    public int? TargetHeaderRow { get; set; }

    /// <summary>Blank rows to leave before appended data.</summary>
    [Parameter]
    public int BlankRowsBefore { get; set; }

    /// <summary>Allow copied values to replace existing target cells.</summary>
    [Parameter]
    public SwitchParameter OverwriteExistingCells { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var targetWorkbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, targetWorkbook.Document, Path, "Update Excel workbook")) {
            return;
        }

        var targetDocument = targetWorkbook.Document;
        var targetSheet = ExcelWorkbookCommandService.ResolveSheet(this, targetDocument, ParameterSetName, TargetSheet, TargetSheetIndex);
        using var sourceWorkbook = ExcelWorkbookCommandService.ResolveSourceWorkbook(this, targetDocument, SourceDocument, SourcePath, readOnly: true);
        var sourceSheet = sourceWorkbook.Document[SourceSheet];

        var result = targetDocument.MergeWorksheets(targetSheet, sourceSheet, BuildOptions());
        targetWorkbook.SaveIfOwned();
        WriteObject(result);
    }

    private ExcelWorksheetMergeOptions BuildOptions() {
        return new ExcelWorksheetMergeOptions {
            SourceRange = SourceRange,
            TargetStartRow = TargetStartRow,
            TargetStartColumn = TargetStartColumn,
            SourceHasHeader = !NoSourceHeader.IsPresent,
            IncludeSourceHeader = IncludeSourceHeader.IsPresent,
            MatchColumnsByHeader = MatchColumnsByHeader.IsPresent,
            TargetHeaderRow = TargetHeaderRow,
            BlankRowsBefore = BlankRowsBefore,
            OverwriteExistingCells = OverwriteExistingCells.IsPresent
        };
    }

}
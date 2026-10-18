#pragma warning disable CS1591
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Compares two workbooks by sheets, cells, formulas, styles, tables, comments, names, and worksheet metadata.</summary>
/// <example>
///   <summary>Compare two generated workbooks before publishing a refreshed report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$diff = Compare-OfficeExcelWorkbook -Path .\Expected.xlsx -DifferencePath .\Actual.xlsx -MaxDifferences 500
/// if (-not $diff.AreEqual) {
///     $diff.Differences |
///         Sort-Object Category,SheetName,Address |
///         Format-Table Category,SheetName,Address,Message,LeftValue,RightValue
/// }</code>
///   <para>Uses OfficeIMO's reusable workbook diff engine and includes structural metadata by default, not just visible cell values.</para>
/// </example>
[Cmdlet(VerbsData.Compare, "OfficeExcelWorkbook", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelWorkbookCompare")]
[OutputType(typeof(PSObject))]
public sealed class CompareOfficeExcelWorkbookCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "ReferencePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook path to compare against.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetPath)]
    public string DifferencePath { get; set; } = string.Empty;

    /// <summary>Workbook document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Workbook document to compare against.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument DifferenceDocument { get; set; } = null!;

    /// <summary>Maximum number of differences to report.</summary>
    [Parameter]
    public int MaxDifferences { get; set; } = 200;
    /// <summary>Skip visible cell value and formula comparison.</summary>
    [Parameter]
    public SwitchParameter SkipCells { get; set; }
    /// <summary>Skip style-index comparison for used cells.</summary>
    [Parameter]
    public SwitchParameter SkipCellStyles { get; set; }
    /// <summary>Skip workbook and sheet-scoped named-range comparison.</summary>
    [Parameter]
    public SwitchParameter SkipNamedRanges { get; set; }
    /// <summary>Skip table metadata comparison.</summary>
    [Parameter]
    public SwitchParameter SkipTables { get; set; }
    /// <summary>Skip worksheet view, validation, and filter metadata comparison.</summary>
    [Parameter]
    public SwitchParameter SkipWorksheetMetadata { get; set; }
    /// <summary>Skip legacy and threaded comment comparison.</summary>
    [Parameter]
    public SwitchParameter SkipComments { get; set; }

    protected override void ProcessRecord() {
        using var left = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        using var right = ParameterSetName == ParameterSetPath
            ? ExcelWorkbookCommandService.OpenWorkbook(this, DifferencePath, readOnly: true)
            : new ExcelWorkbookCommandScope(DifferenceDocument, ownsDocument: false);

        var report = left.Document.CompareWorkbook(right.Document, new ExcelWorkbookDiffOptions {
            MaxDifferences = MaxDifferences,
            CompareCells = !SkipCells.IsPresent,
            CompareCellStyles = !SkipCellStyles.IsPresent,
            CompareNamedRanges = !SkipNamedRanges.IsPresent,
            CompareTables = !SkipTables.IsPresent,
            CompareWorksheetMetadata = !SkipWorksheetMetadata.IsPresent,
            CompareComments = !SkipComments.IsPresent
        });
        var output = new PSObject();
        output.Properties.Add(new PSNoteProperty("AreEqual", report.AreEqual));
        output.Properties.Add(new PSNoteProperty("DifferenceCount", report.Differences.Count));
        output.Properties.Add(new PSNoteProperty("Differences", report.Differences.Select(CreateDifference).ToArray()));
        WriteObject(output);
    }

    private static PSObject CreateDifference(ExcelWorkbookDifference difference) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("Category", difference.Category));
        item.Properties.Add(new PSNoteProperty("Message", difference.Message));
        item.Properties.Add(new PSNoteProperty("SheetName", difference.SheetName));
        item.Properties.Add(new PSNoteProperty("Address", difference.Address));
        item.Properties.Add(new PSNoteProperty("LeftValue", difference.LeftValue));
        item.Properties.Add(new PSNoteProperty("RightValue", difference.RightValue));
        return item;
    }
}
#pragma warning restore CS1591
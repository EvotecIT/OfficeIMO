using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds or refreshes a workbook table of contents sheet.</summary>
/// <para>Can run inside the Excel DSL, against an open workbook, or directly against a file path.</para>
/// <example>
///   <summary>Add a TOC sheet and verify workbook structure.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Add-OfficeExcelTableOfContents -Path .\report.xlsx -IncludeNamedRanges -AddBackLinks -PassThru
///     Get-OfficeExcelSummary -Path .\report.xlsx -IncludeSheets |
///         Select-Object -Property SheetCount, NamedRangeCount, Sheets
/// )
/// $proof</code>
///   <para>Creates or refreshes a TOC sheet, adds back links, and reads back workbook navigation metadata.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelTableOfContents", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelTableOfContents")]
[OutputType(typeof(ExcelDocument), typeof(FileInfo))]
public sealed class AddOfficeExcelTableOfContentsCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Path to the workbook to update in place.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Name of the TOC sheet.</summary>
    [Parameter]
    public string SheetName { get; set; } = "TOC";

    /// <summary>Keep the TOC sheet in its current position instead of moving it first.</summary>
    [Parameter]
    public SwitchParameter DoNotPlaceFirst { get; set; }

    /// <summary>Disable internal hyperlinks in the TOC sheet.</summary>
    [Parameter]
    public SwitchParameter NoHyperlinks { get; set; }

    /// <summary>Include named ranges in the TOC.</summary>
    [Parameter]
    public SwitchParameter IncludeNamedRanges { get; set; }

    /// <summary>Include hidden named ranges when listing named ranges.</summary>
    [Parameter]
    public SwitchParameter IncludeHiddenNamedRanges { get; set; }

    /// <summary>Disable formatted TOC styling.</summary>
    [Parameter]
    public SwitchParameter NoStyle { get; set; }

    /// <summary>Add a quick link back to the TOC on each worksheet.</summary>
    [Parameter]
    public SwitchParameter AddBackLinks { get; set; }

    /// <summary>Row for the back link when <see cref="AddBackLinks"/> is used.</summary>
    [Parameter]
    public int BackLinkRow { get; set; } = 2;

    /// <summary>Column for the back link when <see cref="AddBackLinks"/> is used.</summary>
    [Parameter]
    public int BackLinkColumn { get; set; } = 1;

    /// <summary>Text used for back links.</summary>
    [Parameter]
    public string BackLinkText { get; set; } = "\u2190 TOC";

    /// <summary>Open the workbook after saving when <see cref="Path"/> is used.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit the updated document or file info.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (BackLinkRow < 1) {
            throw new PSArgumentOutOfRangeException(nameof(BackLinkRow));
        }

        if (BackLinkColumn < 1) {
            throw new PSArgumentOutOfRangeException(nameof(BackLinkColumn));
        }

        if (ParameterSetName == ParameterSetPath) {
            ProcessPath();
            return;
        }

        var document = ParameterSetName == ParameterSetDocument
            ? Document
            : ExcelDslContext.Require(this).Document;

        Apply(document);

        if (PassThru.IsPresent) {
            WriteObject(document);
        }
    }

    private void ProcessPath() {
        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (!File.Exists(resolvedPath)) {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        var fileInfo = new FileInfo(resolvedPath);
        var document = ExcelDocumentService.LoadDocument(resolvedPath, readOnly: false, autoSave: false);
        try {
            Apply(document);
            ExcelDocumentService.SaveDocument(document, Open.IsPresent, resolvedPath);
        } catch {
            ExcelDocumentService.CloseDocument(document);
            throw;
        }

        if (PassThru.IsPresent) {
            WriteObject(fileInfo);
        }
    }

    private void Apply(ExcelDocument document) {
        document.AddTableOfContents(
            sheetName: SheetName,
            placeFirst: !DoNotPlaceFirst.IsPresent,
            withHyperlinks: !NoHyperlinks.IsPresent,
            includeNamedRanges: IncludeNamedRanges.IsPresent,
            includeHiddenNamedRanges: IncludeHiddenNamedRanges.IsPresent,
            styled: !NoStyle.IsPresent);

        if (AddBackLinks.IsPresent) {
            AddBackLinksToSheets(document);
        }
    }

    private void AddBackLinksToSheets(ExcelDocument document) {
        var useExplicitPlacement =
            MyInvocation.BoundParameters.ContainsKey(nameof(BackLinkRow)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(BackLinkColumn));

        if (useExplicitPlacement) {
            document.AddBackLinksToToc(
                tocSheetName: SheetName,
                row: BackLinkRow,
                col: BackLinkColumn,
                text: BackLinkText);
            return;
        }

        var tocSheet = document[SheetName];
        foreach (var sheet in document.Sheets) {
            if (string.Equals(sheet.Name, tocSheet.Name, System.StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            var usedRange = sheet.UsedRangeA1;
            var (startRow, _, endRow, _) = A1.ParseRange(usedRange);
            var row = System.Math.Max(1, System.Math.Max(startRow, endRow) + 2);
            sheet.SetInternalLink(row, 1, tocSheet, "A1", BackLinkText);
        }
    }
}
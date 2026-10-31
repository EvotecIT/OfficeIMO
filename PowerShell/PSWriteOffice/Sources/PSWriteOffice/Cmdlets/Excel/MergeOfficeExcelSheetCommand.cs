using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Appends worksheet rows or merges a source range into a target worksheet.</summary>
/// <para>Source accepts a native ExcelDocument, a workbook path or FileInfo. Omit Source to copy rows
/// within the target workbook. File targets are saved in place; live documents remain caller-owned
/// and require an explicit save. The result reports the actual copied rows and columns.</para>
/// <example><summary>Append monthly rows by their column headers.</summary>
/// <code>Merge-OfficeExcelSheet -Path ./Sales.xlsx -Sheet Year -SourceSheet October -MatchColumnsByHeader</code></example>
/// <example><summary>Merge a literal source file at a chosen target cell.</summary>
/// <code>Get-Item -LiteralPath './Combined[1].xlsx' | Merge-OfficeExcelSheet -Sheet Combined -Source (Get-Item -LiteralPath './Branch[1].xlsx') -SourceSheet Sales -Address B5 -HeaderRow 4 -MatchColumnsByHeader</code></example>
[Cmdlet("Merge", "OfficeExcelSheet", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(ExcelWorksheetMergeResult))]
public sealed class MergeOfficeExcelSheetCommand : OfficeCmdlet {
    /// <summary>Gets or sets target workbook paths with filesystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets literal target paths, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets the live target workbook; omission uses the enclosing worksheet scope.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public ExcelDocument? Document { get; set; }
    /// <summary>Gets or sets the target worksheet name, one-based position or native worksheet.</summary>
    [Parameter, OfficeSheetTransformation, ValidateNotNull]
    public OfficeSheetSelector? Sheet { get; set; }
    /// <summary>Gets or sets the source worksheet name, one-based position or native worksheet.</summary>
    [Parameter(Mandatory = true), OfficeSheetTransformation, ValidateNotNull]
    public OfficeSheetSelector SourceSheet { get; set; } = null!;
    /// <summary>Gets or sets another workbook as a native ExcelDocument, path or literal FileInfo.</summary>
    [Parameter, ValidateNotNull]
    public object? Source { get; set; }
    /// <summary>Gets or sets the source A1 range; omission copies its used range.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? SourceRange { get; set; }
    /// <summary>Gets or sets the target start cell; omission appends after the used range.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Address { get; set; }
    /// <summary>Gets or sets whether the first source row is data rather than a header.</summary>
    [Parameter]
    public SwitchParameter NoSourceHeader { get; set; }
    /// <summary>Gets or sets whether to copy the source header along with its data rows.</summary>
    [Parameter]
    public SwitchParameter IncludeSourceHeader { get; set; }
    /// <summary>Gets or sets whether columns are aligned by header text rather than position.</summary>
    [Parameter]
    public SwitchParameter MatchColumnsByHeader { get; set; }
    /// <summary>Gets or sets the one-based target header row used for matching columns.</summary>
    [Parameter, ValidateRange(1, 1048576)]
    public int? HeaderRow { get; set; }
    /// <summary>Gets or sets blank rows before appended data; an explicit Address takes precedence.</summary>
    [Parameter, ValidateRange(0, 1048575)]
    public int BlankRowsBefore { get; set; }
    /// <summary>Gets or sets whether source values may replace existing target cells.</summary>
    [Parameter]
    public SwitchParameter OverwriteExistingCells { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        ExcelWorksheetMergeOptions options;
        try { options = BuildOptions(); }
        catch (ArgumentException error) { InvalidArgument(error, Address); return; }

        if (ParameterSetName == "Document") {
            try {
                ExcelSheet target = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
                if (ExcelShouldProcessService.ShouldProcessWorkbook(this, target.Document, null,
                    "Merge worksheet rows into " + target.Name)) WriteObject(CopyRows(target, options));
            } catch (ArgumentException error) { InvalidArgument(error, Document); }
            catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
                ItemError(error, "excel-sheet-merge-failed", Document);
            }
            return;
        }

        bool literal = ParameterSetName == "LiteralPath";
        foreach (string input in literal ? LiteralPath : Path) {
            IReadOnlyList<string> paths;
            try {
                paths = OfficePath.ResolveInput(SessionState, input, literal,
                    reportItemError: (error, item) => ItemError(error, "input-unreadable", item));
            } catch (ArgumentException error) { InvalidArgument(error, input); return; }
            catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
                ItemError(error, "input-unreadable", input);
                continue;
            }
            foreach (string path in paths) {
                if (!ExcelShouldProcessService.ShouldProcessTarget(this, path, "Merge worksheet rows")) continue;
                try {
                    using var target = new ExcelWorkbookCommandScope(
                        ExcelDocumentService.LoadDocument(path, readOnly: false, autoSave: false), ownsDocument: true);
                    ExcelWorksheetMergeResult result = CopyRows(ExcelSheetResolver.ResolveTarget(this, target.Document, Sheet), options);
                    target.SaveIfOwned();
                    WriteObject(result);
                } catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
                    ItemError(error, "excel-sheet-merge-failed", path);
                }
            }
        }
    }

    private ExcelWorksheetMergeResult CopyRows(ExcelSheet target, ExcelWorksheetMergeOptions options) {
        using var source = ExcelWorkbookCommandService.ResolveSourceWorkbook(this, target.Document, Source);
        ExcelSheet selected = SourceSheet.Resolve(source.Document.Sheets, sheet => sheet.Name);
        return target.Document.MergeWorksheets(target, selected, options);
    }

    private ExcelWorksheetMergeOptions BuildOptions() {
        var options = new ExcelWorksheetMergeOptions {
            SourceRange = SourceRange,
            SourceHasHeader = !NoSourceHeader.IsPresent,
            IncludeSourceHeader = IncludeSourceHeader.IsPresent,
            MatchColumnsByHeader = MatchColumnsByHeader.IsPresent,
            TargetHeaderRow = HeaderRow,
            BlankRowsBefore = BlankRowsBefore,
            OverwriteExistingCells = OverwriteExistingCells.IsPresent
        };
        if (Address != null) {
            var (row, column) = ExcelHostExtensions.ResolveCellAddress(null, null, Address);
            if (row > 1048576 || column > 16384) throw new ArgumentOutOfRangeException(nameof(Address), "The target cell is outside worksheet bounds.");
            options.TargetStartRow = row;
            options.TargetStartColumn = column;
        }
        return options;
    }
}

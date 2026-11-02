using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Copies selected worksheets within a workbook or from another workbook.</summary>
/// <para>Source accepts a native ExcelDocument, a workbook path or literal FileInfo. Sheet accepts names,
/// wildcard patterns, one-based positions and native worksheets. Omit Sheet to copy all source sheets,
/// or the current worksheet inside a worksheet content block. Native package copy preserves references
/// between selected sheets. File targets are saved in place; live documents require an explicit save.</para>
/// <para>PassThru emits native worksheets. For file input, their workbook remains open: close
/// the returned worksheet's Document with Close-OfficeExcel when finished.</para>
/// <example><summary>Copy a template sheet inside a live workbook.</summary>
/// <code>Copy-OfficeExcelSheet -Document $book -Sheet Template -Name October</code></example>
/// <example><summary>Import regional sheets into an existing workbook.</summary>
/// <code>Copy-OfficeExcelSheet -LiteralPath ./Combined.xlsx -Source ./Region.xlsx -Sheet 'Sales*' -NamePrefix 'Region '</code></example>
[Cmdlet(VerbsCommon.Copy, "OfficeExcelSheet", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(ExcelSheet))]
public sealed class CopyOfficeExcelSheetCommand : OfficeCmdlet {
    /// <summary>Gets or sets target workbook paths with filesystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets literal target paths, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets a live target workbook; omission uses the enclosing workbook scope.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public ExcelDocument? Document { get; set; }
    /// <summary>Gets or sets source worksheet names or patterns, one-based positions or native worksheets.</summary>
    [Parameter, OfficeSheetTransformation, SupportsWildcards, ValidateNotNullOrEmpty]
    public OfficeSheetSelector[]? Sheet { get; set; }
    /// <summary>Gets or sets a new name when copying exactly one worksheet.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Name { get; set; }
    /// <summary>Gets or sets a prefix for each copied worksheet name.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? NamePrefix { get; set; }
    /// <summary>Gets or sets another workbook as a native ExcelDocument, path or literal FileInfo.</summary>
    [Parameter, ValidateNotNull]
    public object? Source { get; set; }
    /// <summary>Gets or sets how invalid or duplicate destination names are handled.</summary>
    [Parameter]
    public ExcelSheetNameValidationMode ValidationMode { get; set; } = ExcelSheetNameValidationMode.Sanitize;
    /// <summary>Gets or sets native package copying or explicit value-only materialization.</summary>
    [Parameter]
    public ExcelWorksheetCopyMode CopyMode { get; set; } = ExcelWorksheetCopyMode.Package;
    /// <summary>Gets or sets whether to return live copied worksheets; their workbook is caller-owned.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!Enum.IsDefined(typeof(ExcelSheetNameValidationMode), ValidationMode) ||
            !Enum.IsDefined(typeof(ExcelWorksheetCopyMode), CopyMode)) {
            InvalidArgument(new ArgumentException("Choose a defined ValidationMode and CopyMode."), CopyMode);
            return;
        }
        if (ParameterSetName == "Document") {
            try {
                ExcelDocument target = Document ?? OfficeDocumentContext.Find<ExcelDocument>()
                    ?? ExcelDslContext.Current?.Document ?? ResolveDocument(Document);
                if (ExcelShouldProcessService.ShouldProcessWorkbook(this, target, null, "Copy worksheets"))
                    WriteCopies(CopySheets(target, useCurrentSheet: Document == null));
            } catch (ArgumentException error) { InvalidArgument(error, Document); }
            catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
                ItemError(error, "excel-sheet-copy-failed", Document);
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
                if (!ExcelShouldProcessService.ShouldProcessTarget(this, path, "Copy worksheets")) continue;
                try {
                    using var target = new ExcelWorkbookCommandScope(
                        ExcelDocumentService.LoadDocument(path, readOnly: false, autoSave: false), ownsDocument: true);
                    IReadOnlyList<ExcelSheet> copies = CopySheets(target.Document, useCurrentSheet: false);
                    target.SaveIfOwned();
                    WriteCopies(copies);
                    if (PassThru) target.TransferOwnership();
                } catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
                    ItemError(error, "excel-sheet-copy-failed", path);
                }
            }
        }
    }

    private IReadOnlyList<ExcelSheet> CopySheets(ExcelDocument target, bool useCurrentSheet) {
        using var source = ExcelWorkbookCommandService.ResolveSourceWorkbook(this, target, Source);
        ExcelSheet? current = Source == null && useCurrentSheet && ExcelDslContext.Current?.CurrentSheet is ExcelSheet selected
            && ReferenceEquals(selected.Document, target) ? selected : null;
        IReadOnlyList<ExcelSheet> sheets = ExcelSheetResolver.ResolveMany(source.Document, Sheet, current);
        if (sheets.Count == 0) throw new ArgumentException("No source worksheets matched Sheet.", nameof(Sheet));
        if (Name != null) {
            if (sheets.Count != 1) throw new ArgumentException("Name requires exactly one source worksheet. Use NamePrefix for several sheets.", nameof(Name));
            return new[] { target.CopyWorksheetFrom(source.Document, sheets[0].Name, (NamePrefix ?? string.Empty) + Name,
                ValidationMode, new ExcelWorksheetCopyOptions { CopyMode = CopyMode }) };
        }
        ExcelWorkbookMergeResult result = target.MergeWorkbookFrom(source.Document, new ExcelWorkbookMergeOptions {
            SheetNames = sheets.Select(sheet => sheet.Name).ToArray(),
            SheetNamePrefix = NamePrefix,
            SheetNameValidationMode = ValidationMode,
            CopyMode = CopyMode
        });
        return result.TargetSheets.Select(name => target[name]).ToArray();
    }

    private void WriteCopies(IReadOnlyList<ExcelSheet> copies) {
        if (PassThru) foreach (ExcelSheet copy in copies) WriteObject(copy);
    }
}

using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Merges worksheets from one or more workbooks into a target workbook.</summary>
/// <example>
///   <summary>Merge workbook sheets using the package fast path.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$sources = Get-ChildItem .\Incoming\*.xlsx | Select-Object -ExpandProperty FullName
/// $results = Join-OfficeExcelWorkbook -Path .\Combined.xlsx -SourcePath $sources -CopyMode Package -SheetNamePrefix Import
/// $results | Select-Object SheetCount, SourceSheets, TargetSheets</code>
///   <para>Copies worksheets between packages without importing rows into PowerShell objects, which is the preferred path for large workbook merge workflows.</para>
/// </example>
/// <example>
///   <summary>Import selected source sheets with a prefix.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$merge = Join-OfficeExcelWorkbook -Path .\Target.xlsx -SourcePath .\Source.xlsx -SourceSheet Data,Summary -SheetNamePrefix 'Imported '
/// Get-OfficeExcelSummary -Path .\Target.xlsx |
///     Select-Object Path, WorksheetCount</code>
///   <para>Copies selected worksheets from Source.xlsx into Target.xlsx using OfficeIMO workbook merge logic.</para>
/// </example>
[Cmdlet(VerbsCommon.Join, "OfficeExcelWorkbook", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("Merge-OfficeExcelWorkbook", "ExcelWorkbookJoin", "ExcelWorkbookMerge")]
[OutputType(typeof(ExcelWorkbookMergeResult))]
public sealed class JoinOfficeExcelWorkbookCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Target workbook path to create or update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath", "OutputPath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Target workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Optional source workbook object.</summary>
    [Parameter]
    public ExcelDocument? SourceDocument { get; set; }

    /// <summary>Source workbook paths to merge into the target workbook.</summary>
    [Parameter(Position = 1, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "LiteralPath")]
    public string[]? SourcePath { get; set; }

    /// <summary>Specific source worksheet names to import. Defaults to all source sheets.</summary>
    [Parameter]
    [Alias("SheetName", "Sheet", "WorksheetName")]
    public string[]? SourceSheet { get; set; }

    /// <summary>Prefix added to every imported worksheet name.</summary>
    [Parameter]
    public string? SheetNamePrefix { get; set; }

    /// <summary>Controls how invalid or duplicate destination sheet names are handled.</summary>
    [Parameter]
    public ExcelSheetNameValidationMode ValidationMode { get; set; } = ExcelSheetNameValidationMode.Sanitize;

    /// <summary>Controls whether cross-workbook copies use package-level copy or value materialization.</summary>
    [Parameter]
    public ExcelWorksheetCopyMode CopyMode { get; set; } = ExcelWorksheetCopyMode.Package;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (SourceDocument != null && SourcePath is { Length: > 0 }) {
            throw new PSArgumentException("Specify either -SourceDocument or -SourcePath, not both.");
        }

        if (SourceDocument == null && (SourcePath == null || SourcePath.Length == 0)) {
            throw new PSArgumentException("Provide SourceDocument or SourcePath.");
        }

        if (string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase)) {
            var resolvedTargetPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            var action = File.Exists(resolvedTargetPath)
                ? "Update Excel workbook with merged sheets"
                : "Write Excel workbook with merged sheets";
            if (!ShouldProcess(resolvedTargetPath, action)) {
                return;
            }
        }

        using var targetWorkbook = ResolveTargetWorkbook();
        if (!string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase) &&
            !ExcelShouldProcessService.ShouldProcessWorkbook(this, targetWorkbook.Document, Path, "Merge Excel workbook sheets")) {
            return;
        }

        var results = new List<ExcelWorkbookMergeResult>();

        if (SourceDocument != null) {
            results.Add(MergeSourceWorkbook(targetWorkbook.Document, SourceDocument));
        } else {
            foreach (var sourcePath in SourcePath!) {
                if (string.IsNullOrWhiteSpace(sourcePath)) {
                    continue;
                }

                var resolvedSourcePath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(sourcePath);
                using var sourceDocument = ExcelDocumentService.LoadDocument(resolvedSourcePath, readOnly: true, autoSave: false);
                results.Add(MergeSourceWorkbook(targetWorkbook.Document, sourceDocument));
            }
        }

        targetWorkbook.SaveIfOwned();
        foreach (var result in results) {
            WriteObject(result);
        }
    }

    private ExcelWorkbookCommandScope ResolveTargetWorkbook() {
        if (!string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase)) {
            return ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        }

        var resolvedTargetPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var targetDirectory = System.IO.Path.GetDirectoryName(resolvedTargetPath);
        if (!string.IsNullOrWhiteSpace(targetDirectory)) {
            Directory.CreateDirectory(targetDirectory);
        }

        var document = File.Exists(resolvedTargetPath)
            ? ExcelDocumentService.LoadDocument(resolvedTargetPath, readOnly: false, autoSave: false)
            : ExcelDocument.Create(resolvedTargetPath, new ExcelCreateOptions {
                PersistenceMode = DocumentPersistenceMode.Explicit
            });

        return new ExcelWorkbookCommandScope(document, ownsDocument: true);
    }

    private ExcelWorkbookMergeResult MergeSourceWorkbook(ExcelDocument targetDocument, ExcelDocument sourceDocument) {
        return targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions {
            SheetNames = SourceSheet,
            SheetNamePrefix = SheetNamePrefix,
            SheetNameValidationMode = ValidationMode,
            CopyMode = CopyMode
        });
    }
}
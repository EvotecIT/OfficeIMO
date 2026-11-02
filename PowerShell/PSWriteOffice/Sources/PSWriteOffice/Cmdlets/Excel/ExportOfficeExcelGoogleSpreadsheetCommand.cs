using System;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Selects a live export, an offline translation plan or a provider-neutral request batch.</summary>
public enum GoogleSheetsExportMode {
    /// <summary>Export through a supplied Google Workspace session after confirmation.</summary>
    Export = 0,
    /// <summary>Return the native translation plan without contacting Google.</summary>
    Plan = 1,
    /// <summary>Return the native provider-neutral request batch without contacting Google.</summary>
    Batch = 2
}

/// <summary>Plans, compiles or exports Excel workbooks to Google Sheets.</summary>
/// <example><summary>Inspect translation without signing in.</summary>
/// <code>Export-OfficeExcelGoogleSpreadsheet -Path .\Report.xlsx -Mode Plan</code></example>
[Cmdlet(VerbsData.Export, "OfficeExcelGoogleSpreadsheet", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(GoogleSheetsTranslationPlan), typeof(GoogleSheetsBatch), typeof(GoogleSpreadsheetReference))]
public sealed class ExportOfficeExcelGoogleSpreadsheetCommand : GoogleExportCmdlet<ExcelDocument> {
    /// <summary>One native options object and property hashtables applied to an independent snapshot.</summary>
    [Parameter]
    public object[]? Options { get; set; }

    /// <summary>Configured Google Workspace session, required for Export mode.</summary>
    [Parameter]
    public GoogleWorkspaceSession? Session { get; set; }

    /// <summary>Export by default, or return an offline Plan or Batch.</summary>
    [Parameter]
    public GoogleSheetsExportMode Mode { get; set; }

    /// <summary>Reject translation warnings or errors before returning a result or mutating Google.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <inheritdoc />
    protected override Task ProcessRecordAsync() {
        if (!Enum.IsDefined(typeof(GoogleSheetsExportMode), Mode)) throw new PSArgumentException("Choose Export, Plan or Batch.");
        if (Mode == GoogleSheetsExportMode.Export && Session == null)
            throw new PSArgumentException("Provide -Session for Export mode, or use -Mode Plan or Batch.");
        var options = GoogleExportOptionBinding.Capture((EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext"), Options,
            () => new GoogleSheetsSaveOptions(), value => value.Clone());
        return WithInputsAsync(path => ExcelDocument.Load(path, new ExcelLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit
        }), (document, path) => ExportAsync(document, path, options));
    }

    private async Task ExportAsync(ExcelDocument document, string? path, GoogleSheetsSaveOptions options) {
        CancelToken.ThrowIfCancellationRequested();
        if (Mode == GoogleSheetsExportMode.Plan) {
            var plan = document.BuildGoogleSheetsPlan(options);
            EnsureNoLoss(plan.Report);
            WriteObject(plan);
            return;
        }
        if (Mode == GoogleSheetsExportMode.Batch) {
            var batch = document.BuildGoogleSheetsBatch(options);
            EnsureNoLoss(batch.Report);
            WriteObject(batch);
            return;
        }
        if (FailOnLoss.IsPresent) EnsureNoLoss(document.BuildGoogleSheetsPlan(options).Report);
        if (!ShouldProcess(options.Title ?? path ?? "Google spreadsheet", "Export Excel workbook to Google Sheets")) return;
        WriteObject(await document.ExportToGoogleSheetsAsync(Session!, options, CancelToken).ConfigureAwait(false));
    }

    private void EnsureNoLoss(TranslationReport report) {
        if (FailOnLoss.IsPresent && (report.HasWarnings || report.HasErrors))
            throw new InvalidOperationException("Excel-to-Google Sheets translation reported one or more fidelity losses.");
    }
}

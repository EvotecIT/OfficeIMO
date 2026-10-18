using System;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Plans, compiles, or exports an Excel workbook to Google Sheets.</summary>
[Cmdlet(VerbsData.Export, "OfficeExcelGoogleSpreadsheet", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[OutputType(typeof(GoogleSheetsTranslationPlan), typeof(GoogleSheetsBatch), typeof(GoogleSpreadsheetReference))]
public sealed class ExportOfficeExcelGoogleSpreadsheetCommand : AsyncPSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to an Excel workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Excel workbook to export.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Google Sheets translation and destination settings.</summary>
    [Parameter]
    public GoogleSheetsSaveOptions? Options { get; set; }

    /// <summary>Configured Google Workspace session used for a live export.</summary>
    [Parameter]
    public GoogleWorkspaceSession? Session { get; set; }

    /// <summary>Return the translation plan without compiling requests or contacting Google.</summary>
    [Parameter]
    public SwitchParameter PlanOnly { get; set; }

    /// <summary>Return the provider-neutral request batch without contacting Google.</summary>
    [Parameter]
    public SwitchParameter AsBatch { get; set; }

    /// <summary>Throw when translation reports a warning or error.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        if (PlanOnly.IsPresent && AsBatch.IsPresent) throw new ArgumentException("Use either -PlanOnly or -AsBatch, not both.");
        ExcelDocument? owned = null;
        try
        {
            var document = ParameterSetName == ParameterSetPath
                ? owned = ExcelDocument.Load(SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path))
                : Document;
            if (PlanOnly.IsPresent)
            {
                var plan = document.BuildGoogleSheetsPlan(Options);
                EnsureNoLoss(plan.Report);
                WriteObject(plan);
                return;
            }
            if (AsBatch.IsPresent)
            {
                var batch = document.BuildGoogleSheetsBatch(Options);
                EnsureNoLoss(batch.Report);
                WriteObject(batch);
                return;
            }
            if (Session == null) throw new PSInvalidOperationException("Provide -Session for live export, or use -PlanOnly or -AsBatch.");
            if (FailOnLoss.IsPresent) EnsureNoLoss(document.BuildGoogleSheetsPlan(Options).Report);
            var target = Options?.Title ?? (ParameterSetName == ParameterSetPath ? Path : "Google spreadsheet");
            if (!ShouldProcess(target, "Export Excel workbook to Google Sheets")) return;
            var result = await document.ExportToGoogleSheetsAsync(Session, Options, CancelToken).ConfigureAwait(false);
            WriteObject(result);
        }
        finally
        {
            owned?.Dispose();
        }
    }

    private void EnsureNoLoss(TranslationReport report)
    {
        if (FailOnLoss.IsPresent && (report.HasWarnings || report.HasErrors))
            throw new InvalidOperationException("Excel-to-Google Sheets translation reported one or more fidelity losses.");
    }
}

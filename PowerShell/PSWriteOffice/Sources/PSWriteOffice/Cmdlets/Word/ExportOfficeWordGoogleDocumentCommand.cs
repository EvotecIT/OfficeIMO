using System;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;
using OfficeIMO.GoogleWorkspace;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Selects a live export, an offline translation plan or a provider-neutral request batch.</summary>
public enum GoogleDocsExportMode {
    /// <summary>Export through a supplied Google Workspace session after confirmation.</summary>
    Export = 0,
    /// <summary>Return the native translation plan without contacting Google.</summary>
    Plan = 1,
    /// <summary>Return the native provider-neutral request batch without contacting Google.</summary>
    Batch = 2
}

/// <summary>Plans, compiles or exports Word documents to Google Docs.</summary>
/// <example><summary>Inspect translation without signing in.</summary>
/// <code>Export-OfficeWordGoogleDocument -Path .\Report.docx -Mode Plan</code></example>
[Cmdlet(VerbsData.Export, "OfficeWordGoogleDocument", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(GoogleDocsTranslationPlan), typeof(GoogleDocsBatch), typeof(GoogleDocumentReference))]
public sealed class ExportOfficeWordGoogleDocumentCommand : GoogleExportCmdlet<WordDocument> {
    /// <summary>One native options object and property hashtables applied to an independent snapshot.</summary>
    [Parameter]
    public object[]? Options { get; set; }

    /// <summary>Configured Google Workspace session, required for Export mode.</summary>
    [Parameter]
    public GoogleWorkspaceSession? Session { get; set; }

    /// <summary>Export by default, or return an offline Plan or Batch.</summary>
    [Parameter]
    public GoogleDocsExportMode Mode { get; set; }

    /// <summary>Reject translation warnings or errors before returning a result or mutating Google.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <inheritdoc />
    protected override Task ProcessRecordAsync() {
        if (!Enum.IsDefined(typeof(GoogleDocsExportMode), Mode)) throw new PSArgumentException("Choose Export, Plan or Batch.");
        if (Mode == GoogleDocsExportMode.Export && Session == null)
            throw new PSArgumentException("Provide -Session for Export mode, or use -Mode Plan or Batch.");
        var options = GoogleExportOptionBinding.Capture((EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext"), Options,
            () => new GoogleDocsSaveOptions(), value => value.Clone());
        return WithInputsAsync(path => WordDocument.Load(path, new WordLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit
        }), (document, path) => ExportAsync(document, path, options));
    }

    private async Task ExportAsync(WordDocument document, string? path, GoogleDocsSaveOptions options) {
        CancelToken.ThrowIfCancellationRequested();
        if (Mode == GoogleDocsExportMode.Plan) {
            var plan = document.BuildGoogleDocsPlan(options);
            EnsureNoLoss(plan.Report);
            WriteObject(plan);
            return;
        }
        if (Mode == GoogleDocsExportMode.Batch) {
            var batch = document.BuildGoogleDocsBatch(options);
            EnsureNoLoss(batch.Report);
            WriteObject(batch);
            return;
        }
        if (FailOnLoss.IsPresent) EnsureNoLoss(document.BuildGoogleDocsPlan(options).Report);
        if (!ShouldProcess(options.Title ?? path ?? "Google document", "Export Word document to Google Docs")) return;
        WriteObject(await document.ExportToGoogleDocsAsync(Session!, options, CancelToken).ConfigureAwait(false));
    }

    private void EnsureNoLoss(TranslationReport report) {
        if (FailOnLoss.IsPresent && (report.HasWarnings || report.HasErrors))
            throw new InvalidOperationException("Word-to-Google Docs translation reported one or more fidelity losses.");
    }
}

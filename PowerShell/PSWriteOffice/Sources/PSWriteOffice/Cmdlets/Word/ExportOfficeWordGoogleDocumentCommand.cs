using System;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Plans, compiles, or exports a Word document to Google Docs.</summary>
[Cmdlet(VerbsData.Export, "OfficeWordGoogleDocument", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[OutputType(typeof(GoogleDocsTranslationPlan), typeof(GoogleDocsBatch), typeof(GoogleDocumentReference))]
public sealed class ExportOfficeWordGoogleDocumentCommand : AsyncPSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to a Word document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Word document to export.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Google Docs translation and destination settings.</summary>
    [Parameter]
    public GoogleDocsSaveOptions? Options { get; set; }

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
        WordDocument? owned = null;
        try
        {
            var document = ParameterSetName == ParameterSetPath
                ? owned = WordDocument.Load(SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path))
                : Document;
            if (PlanOnly.IsPresent)
            {
                var plan = document.BuildGoogleDocsPlan(Options);
                EnsureNoLoss(plan.Report);
                WriteObject(plan);
                return;
            }
            if (AsBatch.IsPresent)
            {
                var batch = document.BuildGoogleDocsBatch(Options);
                EnsureNoLoss(batch.Report);
                WriteObject(batch);
                return;
            }
            if (Session == null) throw new PSInvalidOperationException("Provide -Session for live export, or use -PlanOnly or -AsBatch.");
            if (FailOnLoss.IsPresent) EnsureNoLoss(document.BuildGoogleDocsPlan(Options).Report);
            var target = Options?.Title ?? (ParameterSetName == ParameterSetPath ? Path : "Google document");
            if (!ShouldProcess(target, "Export Word document to Google Docs")) return;
            var result = await document.ExportToGoogleDocsAsync(Session, Options, CancelToken).ConfigureAwait(false);
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
            throw new InvalidOperationException("Word-to-Google Docs translation reported one or more fidelity losses.");
    }
}

using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading.Tasks;

namespace PSWriteOffice.Foundation;

/// <summary>Resolves Google export inputs while preserving caller-owned document scopes.</summary>
public abstract class GoogleExportCmdlet<TDocument> : AsyncPSCmdlet where TDocument : class, IDisposable {
    /// <summary>Source document paths, including filesystem wildcard patterns.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal source paths; FileInfo.FullName and provider PSPath bind here.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned document; omission uses the enclosing content block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public TDocument? Document { get; set; }

    /// <summary>Runs exports after resolving provider paths or the current pipeline-thread document scope.</summary>
    protected async Task WithInputsAsync(Func<string, TDocument> load, Func<TDocument, string?, Task> export) {
        if (ParameterSetName == "Document") {
            TDocument? document = Document ?? OfficeDocumentContext.Find<TDocument>();
            if (document == null) {
                ThrowTerminatingError(OfficeError.Create(new ArgumentException(
                    "Supply -Document, pipe a live document, or invoke this command inside a New-/Open- content block."),
                    "document-required", Document, ErrorCategory.InvalidArgument));
                return;
            }
            await export(document, null).ConfigureAwait(false);
            return;
        }
        bool literal = ParameterSetName == "LiteralPath";
        IReadOnlyList<string> paths = OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal);
        foreach (string path in paths) {
            CancelToken.ThrowIfCancellationRequested();
            TDocument document;
            try {
                document = load(path);
            } catch (Exception error) when (IsItemFailure(error)) {
                WriteError(OfficeError.Create(error, "google-export-input-failed", path));
                continue;
            }
            using (document) {
                try {
                    await export(document, path).ConfigureAwait(false);
                } catch (Exception error) when (error is not ArgumentException && IsItemFailure(error)) {
                    WriteError(OfficeError.Create(error, "google-export-failed", path));
                }
            }
        }
    }

    private static bool IsItemFailure(Exception error) => error is not PipelineStoppedException and not ActionPreferenceStopException
        and not OperationCanceledException and not OutOfMemoryException and not StackOverflowException;
}

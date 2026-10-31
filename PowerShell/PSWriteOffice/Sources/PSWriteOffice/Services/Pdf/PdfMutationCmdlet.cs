using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Net;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Pdf;

/// <summary>Binds native PDF edits to an implicit document or atomically published file outputs.</summary>
/// <remarks>Provider resolution, output reservations and confirmation run on the pipeline thread. Native file work uses the existing bounded batch bridge.</remarks>
public abstract class PdfMutationCmdlet : OfficeInputCmdlet {
    private readonly HashSet<string> _outputs = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Explicit native input; omission uses the enclosing New-/Open- content block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public PdfDocument? Document { get; set; }
    /// <summary>Single-file destination. Cannot be combined with OutputDirectory.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateNotNullOrEmpty]
    public string? OutputPath { get; set; }
    /// <summary>Destination folder for one or more inputs, retaining their file names.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateNotNullOrEmpty]
    public string? OutputDirectory { get; set; }
    /// <summary>Refuse replacement of an existing file output.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter NoClobber { get; set; }
    /// <summary>Source password as secure input or a warned plaintext value.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>After authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>Native source parsing, authentication and resource-limit settings.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>Explicitly accept fidelity loss when saving file outputs.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }
    /// <summary>Emit the updated native document or each saved FileInfo.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Gets the operation description used for confirmation and progress.</summary>
    protected abstract string Operation { get; }
    /// <summary>Performs a native edit while preserving its input. Mutable creation metadata may return the input.</summary>
    protected abstract PdfDocument Apply(PdfDocument document, CancellationToken cancellationToken);

    /// <inheritdoc />
    protected override Task ProcessRecordAsync() {
        if (ParameterSetName == "Document") {
            var source = Document ?? OfficeDocumentContext.Find<PdfDocument>();
            if (source == null) throw new PSArgumentException("Supply Document, pipe a live PDF, or use a New-/Open- content block.");
            if (!ShouldProcess("PDF document", Operation)) return Task.CompletedTask;
            PdfDocument result = Apply(source, CancelToken);
            bool adopted = OfficeDocumentContext.Replace(source, result);
            if (PassThru) WriteObject(result);
            else if (!adopted && !ReferenceEquals(source, result)) result.Dispose();
            return Task.CompletedTask;
        }
        if ((OutputPath == null) == (OutputDirectory == null))
            throw new PSArgumentException("Specify exactly one of OutputPath or OutputDirectory.");
        var inputs = ResolveInputs();
        if (OutputPath != null && inputs.Count != 1)
            throw new PSArgumentException("OutputPath requires exactly one input PDF.");
        string? directory = OutputDirectory == null ? null : OfficePath.ResolveOutput(SessionState, OutputDirectory, directory: true);
        var work = new List<(string Input, string Output)>();
        foreach (string input in inputs) {
            string output = directory == null
                ? OfficePath.ResolveOutput(SessionState, OutputPath!, directory: false)
                : System.IO.Path.Combine(directory, System.IO.Path.GetFileName(input));
            if (directory != null) {
                string name = System.IO.Path.GetFileNameWithoutExtension(output), extension = System.IO.Path.GetExtension(output);
                for (int suffix = 2; !_outputs.Add(output); suffix++)
                    output = System.IO.Path.Combine(directory, OfficeFileOutput.CreateCollisionCandidateName(name, extension, suffix));
            }
            if (directory == null && !_outputs.Add(output))
                throw new PSArgumentException("OutputPath takes only one input per invocation; use OutputDirectory for a pipeline.");
            try { OfficePath.ValidateDestination(output, NoClobber.IsPresent); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                WriteError(OfficeError.Create(error, "output-unwritable", output));
                continue;
            }
            if (ShouldProcess(output, Operation)) work.Add((input, output));
        }
        var readOptions = PdfCommandUtilities.CreateReadOptions(LoadOptions,
            Password == null ? null : new NetworkCredential(string.Empty, Password).Password,
            IgnorePermissionRestrictions.IsPresent);
        bool passThru = PassThru.IsPresent;
        var saveOptions = new PdfSaveOptions {
            AllowLoss = AllowLoss.IsPresent,
            ConflictPolicy = NoClobber ? OfficeConversionFileConflictPolicy.FailIfExists : OfficeConversionFileConflictPolicy.Replace
        };
        return RunBatchAsync(work, (item, stop) => {
            using var source = PdfDocument.Load(item.Input, readOptions, stop);
            PdfDocument result = Apply(source, stop);
            try {
                result.Save(item.Output, saveOptions, stop);
                return Task.FromResult<object?>(passThru ? new FileInfo(item.Output) : null);
            } finally {
                if (!ReferenceEquals(source, result)) result.Dispose();
            }
        }, Operation);
    }
}

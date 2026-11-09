using System.Management.Automation;
using System.Net;
using System.Security;
using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Pdf;

/// <summary>Binds native PDF edits to an implicit document or atomically published file outputs.</summary>
/// <remarks>Provider resolution, output reservations and confirmation run on the pipeline thread. Native file work uses the existing bounded batch bridge.</remarks>
public abstract class PdfMutationCmdlet : AsyncPSCmdlet {
    private readonly HashSet<string> _outputs = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Explicit native input; omission uses the enclosing New-/Open- content block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public PdfDocument? Document { get; set; }
    /// <summary>Input files with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PathDirectory")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Literal input files, including FileInfo.FullName.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathDirectory")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Single-file destination.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Path"), Parameter(Mandatory = true, ParameterSetName = "LiteralPath")]
    [ValidateNotNullOrEmpty]
    public string? OutputPath { get; set; }
    /// <summary>Destination folder retaining input file names.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "PathDirectory"), Parameter(Mandatory = true, ParameterSetName = "LiteralPathDirectory")]
    [ValidateNotNullOrEmpty]
    public string? OutputDirectory { get; set; }
    /// <summary>Source password.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "PathDirectory"), Parameter(ParameterSetName = "LiteralPath"), Parameter(ParameterSetName = "LiteralPathDirectory")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>After authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "PathDirectory"), Parameter(ParameterSetName = "LiteralPath"), Parameter(ParameterSetName = "LiteralPathDirectory")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>Native source parsing, authentication and resource settings.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "PathDirectory"), Parameter(ParameterSetName = "LiteralPath"), Parameter(ParameterSetName = "LiteralPathDirectory")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>Emit the updated native document or each saved FileInfo.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "PathDirectory"), Parameter(ParameterSetName = "LiteralPath"), Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Gets the operation description used for confirmation and progress.</summary>
    protected abstract string Operation { get; }
    /// <summary>Performs a native edit while preserving its input. Mutable creation metadata may return the input.</summary>
    protected abstract PdfDocument Apply(PdfDocument document, CancellationToken cancellationToken);

    /// <summary>Validates and snapshots operation settings on the pipeline thread.</summary>
    protected virtual void Prepare() { }

    /// <inheritdoc />
    protected override Task ProcessRecordAsync() {
        Prepare();
        if (ParameterSetName == "Document") {
            var source = Document ?? OfficeDocumentContext.Find<PdfDocument>();
            if (source == null) throw new PSArgumentException("Supply Document, pipe a live PDF, or use a New-/Open- content block.");
            if (!PassThru.IsPresent && !ReferenceEquals(OfficeDocumentContext.Find<PdfDocument>(), source) && source.Pipeline.Input != null)
                throw new PSArgumentException("Use -PassThru to retain the edited PDF, or invoke the operation inside Open-OfficePdf.");
            if (!ShouldProcess("PDF document", Operation)) return Task.CompletedTask;
            PdfDocument result = Apply(source, CancelToken);
            bool adopted = OfficeDocumentContext.Replace(source, result);
            if (PassThru) WriteObject(result);
            else if (!adopted && !ReferenceEquals(source, result)) result.Dispose();
            return Task.CompletedTask;
        }
        if ((OutputPath == null) == (OutputDirectory == null))
            throw new PSArgumentException("Specify exactly one of OutputPath or OutputDirectory.");
        var inputs = OfficePath.ResolveInputs(this, ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal) ? LiteralPath : Path,
            ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal));
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
            try { OfficePath.ValidateDestination(output, refuseExisting: false); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                WriteError(OfficeError.Create(error, "output-unwritable", output));
                continue;
            }
            if (ShouldProcess(output, Operation)) work.Add((input, output));
        }
        var readOptions = PdfCommandUtilities.CreateReadOptions(LoadOptions,
            Password,
            IgnorePermissionRestrictions.IsPresent);
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(IgnorePermissionRestrictions)) is bool ignore)
            readOptions = (readOptions ?? new PdfLoadOptions()).Snapshot(null, null, ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        bool passThru = PassThru.IsPresent;
        var saveOptions = new PdfSaveOptions {
            ConflictPolicy = OfficeConversionFileConflictPolicy.Replace
        };
        return OfficeBatch.RunAsync(work, (item, stop) => {
            using var source = PdfDocument.Load(item.Input, readOptions, stop);
            PdfDocument result = Apply(source, stop);
            try {
                result.Save(item.Output, saveOptions, stop);
                return Task.FromResult<object?>(passThru ? new FileInfo(item.Output) : null);
            } finally {
                if (!ReferenceEquals(source, result)) result.Dispose();
            }
        }, 4, CancelToken, this, Operation, "pdf-mutation-failed");
    }
}

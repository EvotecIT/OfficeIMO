using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Conversion;

public sealed partial class ConvertOfficeDocumentCommand {
    private static readonly UTF8Encoding TextEncoding = new(false, true);

    /// <summary>Literal source content, interpreted according to From rather than as a filename.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Text")]
    [AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Source format family or exact codec extension. Overrides filenames while retaining their physical source identity.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Document")]
    [ArgumentCompleter(typeof(OfficeFormatCompleter))]
    public string? From { get; set; }

    /// <summary>Return converted text to the pipeline instead of publishing a main output file.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter AsString { get; set; }

    private async Task ConvertTextAsync() {
        if (NoClobber && AutoRename) throw new PSArgumentException("NoClobber and AutoRename cannot be combined.");
        string sourceExtension = OfficeFormatSelection.Parse(From)!.Extension;
        string targetExtension = OfficeFormatSelection.Parse(To)!.Extension;
        if (!IsTextExtension(sourceExtension)) throw new PSArgumentException("Text requires a text source format, such as Markdown, Html or Text.");
        bool outputText = AsString.IsPresent || OutputPath == null;
        if (outputText && !IsTextExtension(targetExtension)) throw new PSArgumentException("A binary target requires OutputPath.");
        if (AsString && OutputPath != null) throw new PSArgumentException("AsString and OutputPath cannot be combined.");
        string? outputPath = OutputPath == null ? null : OfficePath.ResolveOutput(SessionState, OutputPath, directory: false);
        if (outputPath != null && !OfficeFormatSelection.Parse(To)!.Matches(System.IO.Path.GetExtension(outputPath)))
            throw new PSArgumentException("To and the OutputPath extension must name the same target format.");
        string sourceName = "source" + sourceExtension;
        var options = CaptureOptions(new[] { sourceName }, targetExtension);
        options.SourceName = sourceName;
        options.TargetName = outputPath == null ? "output" + targetExtension : System.IO.Path.GetFileName(outputPath);
        if (TextEncoding.GetByteCount(Text) > options.MaximumInputBytes)
            throw new PSArgumentException("The literal source exceeds MaximumInputBytes.");
        if (!ShouldProcess(outputPath ?? "in-memory " + targetExtension + " output", "Convert text")) return;
        using var input = new MemoryStream(TextEncoding.GetBytes(Text), writable: false);
        if (outputText) {
            using var output = new MemoryStream();
            OfficeConversionReport report = await OfficeConverter.ConvertAsync(input, OfficeConverter.FormatForExtension(sourceExtension),
                output, OfficeConverter.FormatForExtension(targetExtension), options, CancelToken);
            SetSummary(report);
            WriteObject(TextEncoding.GetString(output.ToArray()));
        } else {
            OfficeConversionReport report = await OfficeConverter.ConvertAsync(input, OfficeConverter.FormatForExtension(sourceExtension),
                outputPath!, options, CancelToken);
            SetSummary(report);
            WriteObject(report);
        }
    }

    private async Task ConvertFilesToTextAsync() {
        if (OutputDirectory != null || SaveJob != null || Resume || RetryFailed || CheckpointPath != null || ReportPath != null || Plan)
            throw new PSArgumentException("AsString cannot be combined with file publication, saved jobs, reports or checkpoints.");
        if (_inputs.Any(Directory.Exists)) throw new PSArgumentException("Select individual files for AsString.");
        string target = OfficeFormatSelection.Parse(To)!.Extension;
        if (!IsTextExtension(target)) throw new PSArgumentException("AsString requires a text target format.");
        var options = CaptureOptions(_inputs, target);
        options.TargetName = "output" + target;
        var errors = new ConcurrentQueue<ErrorRecord>();
        long completed = 0, failed = 0;
        // Confirm before any asynchronous worker starts; workers only use captured paths, settings and the native engine.
        if (!ShouldProcess("in-memory " + target + " output", "Convert documents to text")) return;
        await OfficeBatch.RunAsync(_inputs, async (input, token) => {
            try {
                using var output = new MemoryStream();
                await OfficeConverter.ConvertAsync(input, output, OfficeConverter.FormatForExtension(target), options, token);
                string text = TextEncoding.GetString(output.ToArray());
                Interlocked.Increment(ref completed);
                return (object?)text;
            } catch (Exception error) when (error is not OperationCanceledException and not OutOfMemoryException and not StackOverflowException) {
                Interlocked.Increment(ref failed);
                errors.Enqueue(OfficeError.Create(error, "conversion-failed", input));
                return null;
            }
        }, ThrottleLimit, CancelToken, this, "Convert documents to text", "conversion-failed");
        SetSummary(new OfficeBatchSummary(_inputs.Count, completed, 0, failed, CancelToken.IsCancellationRequested));
        foreach (var error in errors) WriteError(error);
        if (_inputs.Count == 0 && !AllowEmpty)
            WriteError(new ErrorRecord(new InvalidOperationException("No files were selected for text conversion."),
                "OfficeDocumentBatchEmpty", ErrorCategory.InvalidResult, null));
    }

    private static bool IsTextExtension(string extension) => extension.ToLowerInvariant() is
        ".txt" or ".md" or ".markdown" or ".html" or ".htm" or ".rtf" or ".csv" or ".tsv" or ".json" or ".xml" or
        ".adoc" or ".asciidoc" or ".tex" or ".latex" or ".confluence";
}

using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.Text;

namespace OfficeIMO.Tool.Commands.Convert;

internal static class OfficePdfCommand {
    internal const string Usage = """
OfficeIMO.Tool - Office to PDF

Usage:
  officeimo convert <input.docx|input.xlsx|input.pptx> [--output <file.pdf>] [--force]

The command uses the first-party OfficeIMO Word, Excel, or PowerPoint PDF adapter.
Conversion diagnostics are written to standard error.
""";

    internal static async Task<int> RunAsync(
        string[] args,
        Stream standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default) {
        try {
            OfficePdfArguments parsed = OfficePdfArguments.Parse(args);
            if (parsed.Help) {
                await WriteUtf8Async(standardOutput, Usage + Environment.NewLine, cancellationToken).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }

            cancellationToken.ThrowIfCancellationRequested();
            string inputPath = Path.GetFullPath(parsed.InputPath!);
            string outputPath = Path.GetFullPath(parsed.OutputPath!);
            if (!File.Exists(inputPath)) throw new FileNotFoundException("Input document was not found.", inputPath);

            using var pdf = new MemoryStream();
            PdfSaveResult result = Convert(inputPath, pdf);
            if (!result.Succeeded) {
                foreach (string diagnostic in result.Diagnostics) {
                    await standardError.WriteLineAsync(diagnostic).ConfigureAwait(false);
                }
                return (int)OfficeImoToolExitCode.OutputFailed;
            }

            foreach (PdfConversionWarning warning in result.Warnings) {
                await standardError.WriteLineAsync(
                    warning.Severity + " " + warning.Code + " [" + warning.Source + "]: " + warning.Message).ConfigureAwait(false);
            }
            if (result.HasLoss && result.Warnings.Count == 0) {
                await standardError.WriteLineAsync("Warning SourceContentLoss: the source conversion reported possible content loss.").ConfigureAwait(false);
            }
            if (result.Report.HasErrors) {
                return (int)OfficeImoToolExitCode.OutputFailed;
            }

            cancellationToken.ThrowIfCancellationRequested();
            await CommitOutputAsync(pdf.ToArray(), outputPath, parsed.Force, cancellationToken).ConfigureAwait(false);
            await WriteUtf8Async(
                standardOutput,
                outputPath + Environment.NewLine,
                CancellationToken.None).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        } catch (OfficePdfUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (OperationCanceledException) {
            await standardError.WriteLineAsync("Conversion cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (FileNotFoundException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        } catch (OfficePdfOutputExistsException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        } catch (IOException exception) {
            await standardError.WriteLineAsync("I/O failed: " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.UnsupportedInput;
        } catch (UnauthorizedAccessException exception) {
            await standardError.WriteLineAsync("Access failed: " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        } catch (Exception exception) {
            await standardError.WriteLineAsync("Conversion failed: " + exception.GetType().Name + ": " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }

    private static PdfSaveResult Convert(string inputPath, Stream output) {
        switch (Path.GetExtension(inputPath).ToLowerInvariant()) {
            case ".docx":
                using (WordDocument document = WordDocument.Load(inputPath, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                    return document.TrySaveAsPdf(output);
                }
            case ".xlsx":
                using (ExcelDocument document = ExcelDocument.Load(inputPath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                    return document.TrySaveAsPdf(output);
                }
            case ".pptx":
                using (PowerPointPresentation presentation = PowerPointPresentation.Load(inputPath, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                    return presentation.TrySaveAsPdf(output);
                }
            default:
                throw new OfficePdfUsageException("The convert command supports DOCX, XLSX, and PPTX input.");
        }
    }

    internal static async Task CommitOutputAsync(
        byte[] content,
        string outputPath,
        bool force,
        CancellationToken cancellationToken) {
        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

        string temporaryPath = Path.Combine(
            directory ?? Directory.GetCurrentDirectory(),
            ".officeimo-" + Guid.NewGuid().ToString("N") + ".tmp");
        try {
            await using (var stream = new FileStream(
                temporaryPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                bufferSize: 81920,
                options: FileOptions.Asynchronous | FileOptions.SequentialScan)) {
                await stream.WriteAsync(content.AsMemory(), cancellationToken).ConfigureAwait(false);
                await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
            }

            cancellationToken.ThrowIfCancellationRequested();
            try {
                File.Move(temporaryPath, outputPath, overwrite: force);
            } catch (IOException) when (!force && File.Exists(outputPath)) {
                throw new OfficePdfOutputExistsException(outputPath);
            }
            temporaryPath = string.Empty;
        } finally {
            if (temporaryPath.Length > 0 && File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }

    private static async Task WriteUtf8Async(Stream output, string value, CancellationToken cancellationToken) {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        await output.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
    }
}

internal sealed class OfficePdfOutputExistsException : IOException {
    internal OfficePdfOutputExistsException(string outputPath)
        : base("Output already exists. Use --force to replace it: " + outputPath) { }
}

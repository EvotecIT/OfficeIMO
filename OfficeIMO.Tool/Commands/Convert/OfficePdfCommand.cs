using DocumentFormat.OpenXml.Packaging;
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
                    [--max-input-bytes <bytes>] [--max-output-bytes <bytes>]
                    [--max-characters-in-part <characters>]

The command uses the first-party OfficeIMO Word, Excel, or PowerPoint PDF adapter.
Package structure, Open XML part size, and PDF output are bounded by default.
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
            if (!parsed.Force && File.Exists(outputPath)) throw new OfficePdfOutputExistsException(outputPath);

            string temporaryPath = CreateTemporaryOutputPath(outputPath);
            try {
                PdfSaveResult result;
                await using (var destination = new FileStream(
                    temporaryPath,
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None,
                    bufferSize: 81920,
                    options: FileOptions.Asynchronous | FileOptions.SequentialScan)) {
                    using var boundedOutput = new OfficePdfOutputLimitStream(destination, parsed.MaxOutputBytes);
                    result = Convert(inputPath, boundedOutput, parsed);
                    await boundedOutput.FlushAsync(cancellationToken).ConfigureAwait(false);
                }

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
                await CommitOutputAsync(temporaryPath, outputPath, parsed.Force, cancellationToken).ConfigureAwait(false);
                temporaryPath = string.Empty;
                await WriteUtf8Async(
                    standardOutput,
                    outputPath + Environment.NewLine,
                    CancellationToken.None).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            } finally {
                if (temporaryPath.Length > 0 && File.Exists(temporaryPath)) File.Delete(temporaryPath);
            }
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
        } catch (OfficePdfOutputLimitException exception) {
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

    private static PdfSaveResult Convert(
        string inputPath,
        Stream output,
        OfficePdfArguments arguments) {
        OfficePackageSecurityOptions packageSecurity = OfficePackageSecurityOptions.SecureDefaults;
        packageSecurity.MaxPackageBytes = arguments.MaxInputBytes;
        packageSecurity.MaxXmlCharactersInPart = arguments.MaxCharactersInPart;
        var openSettings = new OpenSettings {
            MaxCharactersInPart = arguments.MaxCharactersInPart
        };

        switch (Path.GetExtension(inputPath).ToLowerInvariant()) {
            case ".docx":
                using (WordDocument document = WordDocument.Load(inputPath, new WordLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly,
                    MaxInputBytes = arguments.MaxInputBytes,
                    PackageSecurity = packageSecurity,
                    OpenSettings = openSettings
                })) {
                    return document.TrySaveAsPdf(output);
                }
            case ".xlsx":
                using (ExcelDocument document = ExcelDocument.Load(inputPath, new ExcelLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly,
                    MaxInputBytes = arguments.MaxInputBytes,
                    PackageSecurity = packageSecurity,
                    OpenSettings = openSettings
                })) {
                    return document.TrySaveAsPdf(output);
                }
            case ".pptx":
                using (PowerPointPresentation presentation = PowerPointPresentation.Load(inputPath, new PowerPointLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly,
                    MaxInputBytes = arguments.MaxInputBytes,
                    PackageSecurity = packageSecurity,
                    OpenSettings = openSettings
                })) {
                    return presentation.TrySaveAsPdf(output);
                }
            default:
                throw new OfficePdfUsageException("The convert command supports DOCX, XLSX, and PPTX input.");
        }
    }

    internal static Task CommitOutputAsync(
        string temporaryPath,
        string outputPath,
        bool force,
        CancellationToken cancellationToken) {
        try {
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

        return Task.CompletedTask;
    }

    private static string CreateTemporaryOutputPath(string outputPath) {
        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        return Path.Combine(
            directory ?? Directory.GetCurrentDirectory(),
            ".officeimo-" + Guid.NewGuid().ToString("N") + ".tmp");
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

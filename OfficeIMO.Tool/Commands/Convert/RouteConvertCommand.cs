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

internal static class RouteConvertCommand {
    internal const string Usage = """
OfficeIMO.Tool - document conversion

Usage:
  officeimo convert <input> [output] [--output <file>] [--force]
                    [--max-input-bytes <bytes>] [--max-output-bytes <bytes>]
                    [--max-characters-in-part <characters>]
                    [--text-encoding <name>] [--tab-size 1..32] [--allow-legacy-loss]
                    [--diagram-layers screen|print] [--require-no-loss] [--text-only]

The input and output extensions select one route from the shared OfficeIMO conversion catalog
('officeimo convert --help' lists them); the default output is a sibling PDF.
Email PDF renders the styled message: header block, HTML or plain body, inline images and the
attachment list. --text-only writes the Reader text projection of an email instead.
Draw conversion keeps source page dimensions and uses print-visible layers by default.
DjVu conversion preserves scan pixels, page sizes and existing searchable text.
--require-no-loss rejects source and output fidelity losses before publishing; it is not
available for DOC, DOCX, TXT, XLSX or PPTX to PDF.
Package structure, Open XML part size, and output size are bounded by default.
Conversion diagnostics are written to standard error.
""";

    // Rendered from the catalog so help never drifts from what the command accepts.
    internal static string RouteList() {
        var builder = new StringBuilder("Routes:").Append(Environment.NewLine);
        foreach (OfficeIMO.Workflows.OfficeWorkflowRoute route in OfficeIMO.Workflows.OfficeWorkflowCatalog.ExecutableRoutes
                     .OrderBy(route => route.Source, StringComparer.Ordinal).ThenBy(route => route.Target, StringComparer.Ordinal)) {
            builder.Append("  ").Append((route.Source + " -> " + route.Target).PadRight(28))
                .Append(string.Join(" ", route.SourceExtensions)).Append(" -> ").Append(route.TargetExtension)
                .Append(Environment.NewLine);
        }
        return builder.ToString();
    }


    internal static async Task<int> RunAsync(
        string[] args,
        Stream standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default) {
        try {
            RouteConvertArguments parsed = RouteConvertArguments.Parse(args);
            if (parsed.Help) {
                await WriteUtf8Async(standardOutput, Usage + Environment.NewLine + RouteList(), cancellationToken).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }

            cancellationToken.ThrowIfCancellationRequested();
            string inputPath = Path.GetFullPath(parsed.InputPath!);
            string outputPath = Path.GetFullPath(parsed.OutputPath!);
            if (!File.Exists(inputPath)) throw new FileNotFoundException("Input document was not found.", inputPath);
            if (!parsed.Force && File.Exists(outputPath)) throw new RouteConvertOutputExistsException(outputPath);

            OfficeIMO.Workflows.OfficeWorkflow.RegisterConversions();
            OfficeConversionReport report = await OfficeConverter.ConvertAsync(inputPath, outputPath,
                CreateOptions(parsed), cancellationToken).ConfigureAwait(false);
            await WriteDiagnosticsAsync(report, standardError).ConfigureAwait(false);
            await WriteUtf8Async(standardOutput, report.OutputPath + Environment.NewLine, CancellationToken.None).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        } catch (RouteConvertUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (OperationCanceledException) {
            await standardError.WriteLineAsync("Conversion cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (FileNotFoundException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        } catch (RouteConvertOutputExistsException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        } catch (Exception exception) when (OfficeIMO.Core.Internal.OfficeOutputLimit.Is(exception)) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        } catch (OfficeConversionException exception) {
            if (exception.Report is OfficeConversionReport report) await WriteDiagnosticsAsync(report, standardError).ConfigureAwait(false);
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            IReadOnlyList<OfficeDiagnostic> findings = (exception.Report as OfficeConversionReport)?.Diagnostics ?? [];
            if (findings.Any(finding => finding.Code == OfficeDiagnosticCodes.ConversionInputUnsupported)) return (int)OfficeImoToolExitCode.UnsupportedInput;
            if (findings.Any(finding => finding.Code == OfficeDiagnosticCodes.ConversionInputNotFound)) return (int)OfficeImoToolExitCode.InputNotFound;
            if (findings.Any(finding => finding.Code == OfficeDiagnosticCodes.ConversionRequestInvalid)) return (int)OfficeImoToolExitCode.Usage;
            if (findings.Any(finding => finding.Code == OfficeDiagnosticCodes.ConversionOutputFailed)) return (int)OfficeImoToolExitCode.OutputFailed;
            return (int)OfficeImoToolExitCode.OperationFailed;
        } catch (InvalidDataException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.UnsupportedInput;
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

    private static OfficeConversionOptions CreateOptions(RouteConvertArguments arguments) {
        OfficePackageSecurityOptions packageSecurity = OfficePackageSecurityOptions.SecureDefaults;
        packageSecurity.MaxPackageBytes = arguments.MaxInputBytes;
        packageSecurity.MaxXmlCharactersInPart = arguments.MaxCharactersInPart;
        var openSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = arguments.MaxCharactersInPart };
        return new OfficeConversionOptions {
            RouteId = arguments.Route!.Id,
            OutputProfile = arguments.TextOnly ? OfficeConversionProfile.TextOnly : OfficeConversionProfile.Faithful,
            FileConflictPolicy = arguments.Force ? OfficePublicationPolicy.Replace : OfficePublicationPolicy.Fail,            FailOnLoss = arguments.FailOnLoss,            MaximumInputBytes = arguments.MaxInputBytes,MaximumOutputBytes = arguments.MaxOutputBytes,            MaximumXmlCharactersInPart = arguments.MaxCharactersInPart,            WordLoad = new WordLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly, MaxInputBytes = arguments.MaxInputBytes,
                PackageSecurity = packageSecurity, OpenSettings = openSettings,
                FailOnLoss = Path.GetExtension(arguments.InputPath!).Equals(".doc", StringComparison.OrdinalIgnoreCase) && !arguments.AllowLegacyLoss
            },            ExcelLoad = new ExcelLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly, MaxInputBytes = arguments.MaxInputBytes,
                PackageSecurity = packageSecurity, OpenSettings = openSettings
            },            PowerPointLoad = new PowerPointLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly, MaxInputBytes = arguments.MaxInputBytes,
                PackageSecurity = packageSecurity, OpenSettings = openSettings
            },            PlainText = new PdfTextImportOptions { EncodingName = arguments.TextEncoding, TabSize = arguments.TabSize },            RegisteredSettings = new OfficeIMO.Workflows.OfficeWorkflowRouteSettings {
                DjVu = arguments.IsDjVuInput ? new OfficeIMO.DjVu.Pdf.DjVuPdfExportOptions() : null,
                Draw = arguments.IsDrawInput ? new OfficeIMO.OpenDocument.Odg.Pdf.OdgPdfExportOptions {
                    ForPrint = arguments.DiagramForPrint, LossPolicy = OfficeIMO.OpenDocument.OdfConversionLossPolicy.ReportOnly
                } : null
            }
        };
    }

    private static async Task WriteDiagnosticsAsync(OfficeConversionReport report, TextWriter error) {
        foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
            string context = diagnostic.Format is null ? string.Empty : " [" + diagnostic.Format + "]";
            if (diagnostic.Location is not null) context += " [" + diagnostic.Location + "]";
            await error.WriteLineAsync(diagnostic.Severity + " " + diagnostic.Code + context + ": " + diagnostic.Message).ConfigureAwait(false);
        }
    }

    private static async Task WriteUtf8Async(Stream output, string value, CancellationToken cancellationToken) {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        await output.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
    }
}

internal sealed class RouteConvertOutputExistsException : IOException {
    internal RouteConvertOutputExistsException(string outputPath)
        : base("Output already exists. Use --force to replace it: " + outputPath) { }
}

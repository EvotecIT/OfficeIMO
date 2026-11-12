using System.Globalization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

internal static class WorkflowPrintCommand {
    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error, CancellationToken token,
        IPdfPrinterService? printerService = null) {
        IPdfPrinterService service = printerService ?? new PdfPrinterService();
        if (args.Skip(1).Any(argument => argument is "--help" or "-h")) {
            await output.WriteLineAsync(WorkflowCommand.Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        if (args[0] == "printers") {
            if (args.Length == 3 && args[1] == "--paper-sources") {
                foreach (PdfPaperSourceInfo source in (await service.GetCapabilitiesAsync(args[2], token).ConfigureAwait(false)).PaperSources)
                    await output.WriteLineAsync(source.Id + ": " + source.Name).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }
            if (args.Length == 3 && args[1] == "--capabilities") {
                await WriteCapabilitiesAsync(await service.GetCapabilitiesAsync(args[2], token).ConfigureAwait(false), output).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }
            if (args.Length != 1) throw new WorkflowUsageException("Use printers, printers --capabilities <printer> or printers --paper-sources <printer>.");
            foreach (PdfPrinterInfo queue in await service.GetPrintersAsync(token).ConfigureAwait(false))
                await output.WriteLineAsync(queue.Name + (queue.IsDefault ? " (default)" : "") + (queue.RequiresOutputFile ? " (file printer)" : "")).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        if (args[0] == "print-job") return await RunJobAsync(args, output, service, token).ConfigureAwait(false);

        var planner = new List<string> { "print-plan" };
        string? printer = null, paperSource = null, outputFile = null;
        int copies = 1;
        double dpi = 300;
        PdfPrintDuplex duplex = PdfPrintDuplex.PrinterDefault;
        for (int index = 1; index < args.Length; index++) {
            string option = args[index];
            if (option is "--printer" or "--copies" or "--duplex" or "--paper-source" or "--dpi" or "--output-file") {
                if (++index >= args.Length) throw new WorkflowUsageException(option + " requires a value.");
                string value = args[index];
                switch (option) {
                    case "--printer": printer = value; break;
                    case "--paper-source": paperSource = value; break;
                    case "--output-file": outputFile = Path.GetFullPath(value); break;
                    case "--copies":
                        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out copies) || copies is < 1 or > 100)
                            throw new WorkflowUsageException("Copies must be between 1 and 100.");
                        break;
                    case "--dpi":
                        if (!double.TryParse(value, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out dpi) || !double.IsFinite(dpi) || dpi is < 72 or > 600)
                            throw new WorkflowUsageException("Printing resolution must be between 72 and 600 DPI.");
                        break;
                    case "--duplex": duplex = value switch {
                        "default" => PdfPrintDuplex.PrinterDefault, "off" => PdfPrintDuplex.SingleSided,
                        "long" => PdfPrintDuplex.LongEdge, "short" => PdfPrintDuplex.ShortEdge,
                        _ => throw new WorkflowUsageException("Duplex must be default, off, long, or short.")
                    }; break;
                }
            } else planner.Add(option);
        }
        if (string.IsNullOrWhiteSpace(printer)) throw new WorkflowUsageException("Printing requires an explicit --printer name or IPP URI.");
        WorkflowArguments parsed = WorkflowArguments.Parse(planner.ToArray());
        if (parsed.Command != WorkflowCommandKind.PrintPlan) throw new WorkflowUsageException("Choose a document and supported print-plan options.");
        PdfPrintPlanRequest layout = parsed.CreatePrintRequest();
        if (layout.PdfPassword is not null && !string.Equals(Path.GetExtension(layout.InputPath), ".pdf", StringComparison.OrdinalIgnoreCase))
            throw new WorkflowUsageException("--password-env applies only to PDF input.");
        try {
            OfficePrintResult result = await service.PrintAsync(new OfficePrintRequest {
                Layout = layout,
                // Without --paper, sheets use the printer's own default media.
                UsePrinterPaper = !parsed.PaperSpecified,
                Delivery = new PdfPrintDeliveryOptions {
                    PrinterName = printer!, DocumentName = Path.GetFileName(layout.InputPath), Copies = copies, Duplex = duplex, PaperSourceId = paperSource,
                    OutputFilePath = outputFile, MaximumRasterDpi = dpi
                }
            }, token).ConfigureAwait(false);
            if (result.Conversion is { } conversion)
                foreach (OfficeDiagnostic diagnostic in conversion.Diagnostics.Where(item => item.Severity != OfficeDiagnosticSeverity.Information))
                    await error.WriteLineAsync("Conversion " + diagnostic.Severity.ToString().ToLowerInvariant() + ": " + diagnostic.Message).ConfigureAwait(false);
            foreach (string diagnostic in result.Diagnostics) await error.WriteLineAsync(diagnostic).ConfigureAwait(false);
            PdfPrintSubmission receipt = result.Submission;
            await output.WriteLineAsync("Queue accepted: " + receipt.PrinterName + "; job " + receipt.JobId + ". Physical delivery is unconfirmed.").ConfigureAwait(false);
            if (receipt.IgnoredAttributes.Count > 0)
                await error.WriteLineAsync("The printer ignored: " + string.Join(", ", receipt.IgnoredAttributes) + ".").ConfigureAwait(false);
            if (receipt.CleanupWarning != null) await error.WriteLineAsync(receipt.CleanupWarning).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        } catch (PdfPrintDeliveryException exception) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        }
    }

    /// <summary>print-job status|wait|cancel --printer &lt;printer&gt; --job &lt;id&gt; [--timeout &lt;seconds&gt;]</summary>
    private static async Task<int> RunJobAsync(string[] args, TextWriter output, IPdfPrinterService service, CancellationToken token) {
        string action = args.Length > 1 ? args[1] : string.Empty;
        if (action is not ("status" or "wait" or "cancel")) throw new WorkflowUsageException("Use print-job status|wait|cancel --printer <printer> --job <id>.");
        string? printer = null, job = null;
        double timeout = 300;
        for (int index = 2; index < args.Length; index++) {
            string option = args[index];
            if (++index >= args.Length) throw new WorkflowUsageException(option + " requires a value.");
            string value = args[index];
            switch (option) {
                case "--printer": printer = value; break;
                case "--job": job = value; break;
                case "--timeout" when action == "wait":
                    if (!double.TryParse(value, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out timeout) || !double.IsFinite(timeout) || timeout is <= 0 or > 86400)
                        throw new WorkflowUsageException("The timeout must be from 1 through 86400 seconds.");
                    break;
                default: throw new WorkflowUsageException("Unknown option '" + option + "'.");
            }
        }
        if (string.IsNullOrWhiteSpace(printer) || string.IsNullOrWhiteSpace(job)) throw new WorkflowUsageException("print-job requires --printer and --job.");
        PdfPrintJobStatus status = action switch {
            "cancel" => await service.CancelJobAsync(printer!, job!, token).ConfigureAwait(false),
            "wait" => await service.WaitForJobAsync(printer!, job!, TimeSpan.FromSeconds(timeout), cancellationToken: token).ConfigureAwait(false),
            _ => await service.GetJobAsync(printer!, job!, token).ConfigureAwait(false)
        };
        await output.WriteLineAsync("Job " + status.JobId + " on " + status.PrinterName + ": " + status.State
            + (status.StateReasons.Count > 0 ? " (" + string.Join(", ", status.StateReasons) + ")" : "")
            + (string.IsNullOrEmpty(status.StateMessage) ? "" : ": " + status.StateMessage)).ConfigureAwait(false);
        return (int)OfficeImoToolExitCode.Success;
    }

    private static async Task WriteCapabilitiesAsync(PdfPrinterCapabilities capabilities, TextWriter output) {
        static string State(bool? value) => value switch { true => "yes", false => "no", _ => "unknown" };
        await output.WriteLineAsync("Printer: " + capabilities.PrinterName).ConfigureAwait(false);
        await output.WriteLineAsync("Default media: " + (capabilities.DefaultMedia?.Name ?? "unknown")).ConfigureAwait(false);
        foreach (PdfPrinterMedia media in capabilities.Media) {
            string size = media.Size is { } paper
                ? " (" + paper.Width.ToPoints().ToString("0.##", CultureInfo.InvariantCulture) + " x " + paper.Height.ToPoints().ToString("0.##", CultureInfo.InvariantCulture) + " pt)"
                : string.Empty;
            await output.WriteLineAsync("Media: " + media.Id + ": " + media.Name + size).ConfigureAwait(false);
        }
        foreach (PdfPaperSourceInfo source in capabilities.PaperSources)
            await output.WriteLineAsync("Paper source: " + source.Id + ": " + source.Name).ConfigureAwait(false);
        await output.WriteLineAsync("Duplex: " + State(capabilities.SupportsDuplex)).ConfigureAwait(false);
        await output.WriteLineAsync("Color: " + State(capabilities.SupportsColor)).ConfigureAwait(false);
        await output.WriteLineAsync("Collation: " + State(capabilities.SupportsCollation)).ConfigureAwait(false);
        await output.WriteLineAsync("Maximum copies: " + (capabilities.MaximumCopies?.ToString(CultureInfo.InvariantCulture) ?? "unknown")).ConfigureAwait(false);
        if (capabilities.Resolutions.Count > 0)
            await output.WriteLineAsync("Resolutions: " + string.Join(", ", capabilities.Resolutions)).ConfigureAwait(false);
    }
}

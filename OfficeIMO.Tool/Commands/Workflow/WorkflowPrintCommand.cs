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
            string? queueName = null;
            bool capabilities = false, listJson = true;
            for (int index = 1; index < args.Length; index++) {
                string option = args[index];
                if (option is not ("--paper-sources" or "--capabilities" or "--format") || ++index >= args.Length || args[index].StartsWith("--", StringComparison.Ordinal))
                    throw new WorkflowUsageException("Use printers [--capabilities <printer> | --paper-sources <printer>] [--format json|text].");
                if (option == "--format") {
                    listJson = WorkflowOutput.ParseJson(args[index]);
                    continue;
                }
                if (queueName != null) throw new WorkflowUsageException("Choose either --capabilities or --paper-sources.");
                queueName = args[index];
                capabilities = option == "--capabilities";
            }
            if (queueName == null) {
                await WorkflowOutput.WritePrintersAsync(output, await service.GetPrintersAsync(token).ConfigureAwait(false), listJson).ConfigureAwait(false);
            } else {
                PdfPrinterCapabilities found = await service.GetCapabilitiesAsync(queueName, token).ConfigureAwait(false);
                if (capabilities) await WorkflowOutput.WritePrinterCapabilitiesAsync(output, found, listJson).ConfigureAwait(false);
                else await WorkflowOutput.WritePaperSourcesAsync(output, queueName, found.PaperSources, listJson).ConfigureAwait(false);
            }
            return (int)OfficeImoToolExitCode.Success;
        }
        if (args[0] == "print-job") return await RunJobAsync(args, output, service, token).ConfigureAwait(false);

        var planner = new List<string> { "print-plan" };
        string? printer = null, paperSource = null, outputFile = null;
        int copies = 1;
        bool json = true;
        double dpi = 300;
        PdfPrintDuplex duplex = PdfPrintDuplex.PrinterDefault;
        for (int index = 1; index < args.Length; index++) {
            string option = args[index];
            if (option is "--format" or "--printer" or "--copies" or "--duplex" or "--paper-source" or "--dpi" or "--output-file") {
                if (++index >= args.Length) throw new WorkflowUsageException(option + " requires a value.");
                string value = args[index];
                switch (option) {
                    case "--format": json = WorkflowOutput.ParseJson(value); break;
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
            await WorkflowOutput.WritePrintAsync(output, receipt, layout.InputPath, json).ConfigureAwait(false);
            if (receipt.IgnoredAttributes.Count > 0)
                await error.WriteLineAsync("The printer ignored: " + string.Join(", ", receipt.IgnoredAttributes) + ".").ConfigureAwait(false);
            if (receipt.CleanupWarning != null) await error.WriteLineAsync(receipt.CleanupWarning).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        } catch (PdfPrintDeliveryException exception) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        }
    }

    /// <summary>print-job status|wait|cancel --printer &lt;printer&gt; --job &lt;id&gt; [--timeout &lt;seconds&gt;] [--format json|text]</summary>
    private static async Task<int> RunJobAsync(string[] args, TextWriter output, IPdfPrinterService service, CancellationToken token) {
        string action = args.Length > 1 ? args[1] : string.Empty;
        if (action is not ("status" or "wait" or "cancel")) throw new WorkflowUsageException("Use print-job status|wait|cancel --printer <printer> --job <id>.");
        string? printer = null, job = null;
        double timeout = 300;
        bool json = true;
        for (int index = 2; index < args.Length; index++) {
            string option = args[index];
            if (++index >= args.Length) throw new WorkflowUsageException(option + " requires a value.");
            string value = args[index];
            switch (option) {
                case "--printer": printer = value; break;
                case "--job": job = value; break;
                case "--format": json = WorkflowOutput.ParseJson(value); break;
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
        await WorkflowOutput.WritePrintJobAsync(output, status, json).ConfigureAwait(false);
        return (int)OfficeImoToolExitCode.Success;
    }
}

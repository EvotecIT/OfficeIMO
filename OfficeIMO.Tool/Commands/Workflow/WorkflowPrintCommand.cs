using System.Globalization;
using OfficeIMO.Pdf;
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
                foreach (PdfPaperSourceInfo source in await service.GetPaperSourcesAsync(args[2], token).ConfigureAwait(false))
                    await output.WriteLineAsync(source.Id + ": " + source.Name).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }
            if (args.Length != 1) throw new WorkflowUsageException("Use printers or printers --paper-sources <queue>.");
            foreach (PdfPrinterInfo queue in await service.GetPrintersAsync(token).ConfigureAwait(false))
                await output.WriteLineAsync(queue.Name + (queue.IsDefault ? " (default)" : "") + (queue.RequiresOutputFile ? " (file printer)" : "")).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        var planner = new List<string> { "print-plan" };
        string? printer = null, paperSource = null, outputFile = null;
        int copies = 1;
        double dpi = 150;
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
        if (string.IsNullOrWhiteSpace(printer)) throw new WorkflowUsageException("Printing requires an explicit --printer name.");
        WorkflowArguments parsed = WorkflowArguments.Parse(planner.ToArray());
        if (parsed.Command != WorkflowCommandKind.PrintPlan) throw new WorkflowUsageException("Choose a PDF and supported print-plan options.");
        string input = Path.GetFullPath(parsed.Inputs[0]);
        PdfDocument document = await PdfDocument.LoadAsync(input, cancellationToken: token).ConfigureAwait(false);
        PdfPreparedPrintDocument sheets = PdfPrintRenderer.Prepare(document, new PdfPrintPlanRequest {
            InputPath = input, Pages = parsed.Pages, PaperSize = parsed.PaperSize,
            Orientation = parsed.Orientation, PagesPerSheet = parsed.PagesPerSheet, ScaleMode = parsed.ScaleMode, Margin = parsed.Margin
        }, new PdfPrintRenderOptions { Dpi = dpi }, token);
        foreach (var diagnostic in sheets.Diagnostics) await error.WriteLineAsync(diagnostic).ConfigureAwait(false);
        try {
            PdfPrintSubmission receipt = await service.SubmitAsync(sheets, new PdfPrintDeliveryOptions {
                PrinterName = printer, DocumentName = Path.GetFileName(input), Copies = copies, Duplex = duplex, PaperSourceId = paperSource, OutputFilePath = outputFile
            }, token).ConfigureAwait(false);
            await output.WriteLineAsync("Queue accepted: " + receipt.PrinterName + "; job " + (receipt.JobId ?? "unknown") + ". Physical delivery is unconfirmed.").ConfigureAwait(false);
            if (receipt.CleanupWarning != null) await error.WriteLineAsync(receipt.CleanupWarning).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        } catch (PdfPrintDeliveryException exception) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        }
    }
}

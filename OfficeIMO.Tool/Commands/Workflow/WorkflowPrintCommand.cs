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
            string? queueName = null;
            bool listJson = true;
            for (int index = 1; index < args.Length; index++) {
                string option = args[index];
                if (option is not ("--paper-sources" or "--format") || ++index >= args.Length || args[index].StartsWith("--", StringComparison.Ordinal))
                    throw new WorkflowUsageException("Use printers [--paper-sources <queue>] [--format json|text].");
                if (option == "--format") listJson = WorkflowOutput.ParseJson(args[index]);
                else queueName = args[index];
            }
            if (queueName != null)
                await WorkflowOutput.WritePaperSourcesAsync(output, queueName, await service.GetPaperSourcesAsync(queueName, token).ConfigureAwait(false), listJson).ConfigureAwait(false);
            else
                await WorkflowOutput.WritePrintersAsync(output, await service.GetPrintersAsync(token).ConfigureAwait(false), listJson).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        var planner = new List<string> { "print-plan" };
        string? printer = null, paperSource = null, outputFile = null;
        int copies = 1;
        bool json = true;
        double dpi = 150;
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
        if (string.IsNullOrWhiteSpace(printer)) throw new WorkflowUsageException("Printing requires an explicit --printer name.");
        WorkflowArguments parsed = WorkflowArguments.Parse(planner.ToArray());
        if (parsed.Command != WorkflowCommandKind.PrintPlan) throw new WorkflowUsageException("Choose a PDF and supported print-plan options.");
        string input = Path.GetFullPath(parsed.Inputs[0]);
        PdfDocument document = await PdfDocument.LoadAsync(input, cancellationToken: token).ConfigureAwait(false);
        PdfPreparedPrintDocument sheets = PdfPrintRenderer.Prepare(document, parsed.CreatePrintRequest(), new PdfPrintRenderOptions { Dpi = dpi }, token);
        foreach (var diagnostic in sheets.Diagnostics) await error.WriteLineAsync(diagnostic).ConfigureAwait(false);
        try {
            PdfPrintSubmission receipt = await service.SubmitAsync(sheets, new PdfPrintDeliveryOptions {
                PrinterName = printer, DocumentName = Path.GetFileName(input), Copies = copies, Duplex = duplex, PaperSourceId = paperSource, OutputFilePath = outputFile
            }, token).ConfigureAwait(false);
            await WorkflowOutput.WritePrintAsync(output, receipt, input, json).ConfigureAwait(false);
            if (receipt.CleanupWarning != null) await error.WriteLineAsync(receipt.CleanupWarning).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        } catch (PdfPrintDeliveryException exception) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        }
    }
}

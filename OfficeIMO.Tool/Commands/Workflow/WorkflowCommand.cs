using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

internal static class WorkflowCommand {
    internal const string Usage = """
OfficeIMO.Tool - output and intake workflows

Usage:
  officeimo workflow compare <expected.pdf> <actual.pdf> --output <report.html>
             [--expected-pages <selection>] [--actual-pages <selection>] [--force]
             [--password-env <name>] [--comparison-password-env <name>]
  officeimo workflow optimize-images <input.docx|input.doc>... [--analyze]
             [--output <file> | --output-directory <folder>] [--format docx|doc|pdf]
             [--mode downsample|recompress|both] [--dpi <36-1200>] [--quality <1-100>] [--force]
  officeimo workflow export-pages <input.pdf> --output <folder> [--pages <selection>]
             [--image-format png|jpeg|webp|tiff|svg] [--dpi <36-600>] [--max-dimension <pixels>] [--force]
             [--format json|text]
  officeimo workflow assemble <source>... --output <output.pdf> [--no-recursive] [--force] [--format json|text]
  officeimo workflow print-plan <input.pdf> [--pages <selection>] [--paper A4|Letter|Legal|A3] [--password-env <name>]
             [--orientation auto|portrait|landscape] [--pages-per-sheet 1|2|4|6|9]
             [--scale fit|actual|fill|custom] [--custom-scale <percent>] [--margin <points>]
             [--margin-left <points>] [--margin-top <points>] [--margin-right <points>] [--margin-bottom <points>]
             [--alignment center|top-left|top|top-right|left|right|bottom-left|bottom|bottom-right]
             [--page-subset all|odd|even] [--color color|grayscale] [--format json|text]
  officeimo workflow batch [<source>...] --output <folder> [--input-directory <folder>]
             [--target pdf|docx|xlsx|pptx|html|json|jsonl|yaml|toml|xml|csv|tsv] [--checkpoint <folder>] [--route <id>]
             [--source-extension <extension>] [--no-recursive] [--concurrency 1..32]
             [--conflict fail|rename|replace] [--retry-failed] [--text-encoding <name>] [--tab-size 1..32]
             [--allow-legacy-loss] [--pages <selection>]
             [--profile faithful|lightweight|printready|textonly] [--maximum-files <count>]
             [--maximum-input-bytes <bytes>] [--maximum-output-bytes <bytes>]
             [--maximum-xml-characters-in-part <characters>] [--diagram-layers screen|print] [--require-no-loss]
             [--source-password-env <name>] [--pdf-password-env <name>] [--allow-empty] [--verbose]
             [--format json|ndjson] [--save-job <new-job.json>]
  officeimo workflow job plan <job.json>
  officeimo workflow job run <job.json> [--retry-failed] [--allow-empty] [--verbose] [--format json|ndjson]
             [--report <new-report.json>]
  officeimo workflow job report <run-report.json>
  officeimo workflow printers [--capabilities <printer> | --paper-sources <printer>]
  officeimo workflow print <document> --printer <queue|ipp://uri> [print-plan options; --paper defaults to the printer's media]
             [--copies 1..100] [--duplex default|off|long|short] [--paper-source <id>] [--dpi 72..600]
             [--output-file <new-local-path>]
  officeimo workflow print-job status|wait|cancel --printer <queue|ipp://uri> --job <id> [--timeout <seconds>]

Page selections accept document-relative expressions such as 1-3,last.
print accepts PDF and every format the conversion catalog turns into PDF; --password-env applies to PDF input.
Printers are installed queue names or ipp:// and ipps:// printer URIs. Queue acceptance is not completion; use print-job.
Folders and ZIP archives are expanded deterministically by the reusable workflow owner.
Existing output is refused unless --force is supplied.
export-pages, assemble and print-plan write one JSON result with a versioned schema identifier;
--format text writes a short summary instead. Diagnostics go to standard error.
Batch output defaults to the existing aggregate JSON. --format json adds per-file outcomes
(up to 10,000 or 16 MiB of items with an explicit truncation marker); ndjson streams every item and a final summary.
Batch and saved jobs fail when no inputs are selected unless --allow-empty is supplied.
--target defaults to pdf, or to the target of --route when a route is named. Data targets (json, jsonl,
yaml, toml, xml, csv, tsv and xlsx) run the structured-data routes.
Files without a route to the target are skipped and counted on standard error; --verbose lists each one.
Passwords are read only from the environment variables named by --source-password-env and --pdf-password-env.
--save-job writes a definition without converting files. Saved jobs contain secret references only.
--report records an attempt independently of checkpoint receipts; job report inspects it without executing work.
""";

    internal static async Task<int> RunAsync(
        string[] args,
        TextWriter standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default,
        IOfficeOutputWorkflowRunner? runner = null,
        Func<PdfPrintPlanRequest, CancellationToken, Task<PdfPrintPlan>>? printPlanner = null,
        IPdfPrinterService? printerService = null) {
        WorkflowCommandKind activeCommand = WorkflowCommandKind.Help;
        try {
            if (args.FirstOrDefault() == "compare")
                return await WorkflowCompareCommand.RunAsync(args[1..], standardOutput, standardError, cancellationToken).ConfigureAwait(false);
            if (args.FirstOrDefault()?.Equals("optimize-images", StringComparison.OrdinalIgnoreCase) == true)
                return await WordImagesCommand.RunAsync(args.Skip(1).ToArray(), standardOutput, standardError, cancellationToken).ConfigureAwait(false);
            if (args.FirstOrDefault() == "batch")
                return await WorkflowBatchCommand.RunAsync(args[1..], standardOutput, standardError, cancellationToken).ConfigureAwait(false);
            if (args.FirstOrDefault() == "job")
                return await WorkflowJobCommand.RunAsync(args[1..], standardOutput, standardError, cancellationToken).ConfigureAwait(false);
            if (args.FirstOrDefault() is "printers" or "print" or "print-job")
                return await WorkflowPrintCommand.RunAsync(args, standardOutput, standardError, cancellationToken, printerService).ConfigureAwait(false);
            WorkflowArguments parsed = WorkflowArguments.Parse(args);
            activeCommand = parsed.Command;
            if (parsed.Command == WorkflowCommandKind.Help) {
                await standardOutput.WriteLineAsync(Usage).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }
            return parsed.Command switch {
                WorkflowCommandKind.ExportPages => await ExportPagesAsync(parsed, runner ?? new OfficeWorkflowRunner(), standardOutput, standardError, cancellationToken).ConfigureAwait(false),
                WorkflowCommandKind.Assemble => await AssembleAsync(parsed, runner ?? new OfficeWorkflowRunner(), standardOutput, standardError, cancellationToken).ConfigureAwait(false),
                WorkflowCommandKind.PrintPlan => await PrintPlanAsync(
                    parsed,
                    standardOutput,
                    printPlanner ?? PdfPrintPlanner.CreateAsync,
                    cancellationToken).ConfigureAwait(false),
                _ => (int)OfficeImoToolExitCode.Usage
            };
        } catch (WorkflowUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (OperationCanceledException) {
            await standardError.WriteLineAsync("Workflow cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (FileNotFoundException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        } catch (UnauthorizedAccessException exception) {
            bool inputOnlyCommand = activeCommand == WorkflowCommandKind.PrintPlan;
            await standardError.WriteLineAsync(
                (inputOnlyCommand ? "Input access failed: " : "Output access failed: ") + exception.Message).ConfigureAwait(false);
            return inputOnlyCommand
                ? (int)OfficeImoToolExitCode.OperationFailed
                : (int)OfficeImoToolExitCode.OutputFailed;
        } catch (IOException exception) {
            bool inputOnlyCommand = activeCommand == WorkflowCommandKind.PrintPlan;
            await standardError.WriteLineAsync(
                (inputOnlyCommand ? "Input I/O failed: " : "Output I/O failed: ") + exception.Message).ConfigureAwait(false);
            return inputOnlyCommand
                ? (int)OfficeImoToolExitCode.OperationFailed
                : (int)OfficeImoToolExitCode.OutputFailed;
        } catch (Exception exception) {
            await standardError.WriteLineAsync("Workflow failed: " + exception.GetType().Name + ": " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }

    private static async Task<int> ExportPagesAsync(
        WorkflowArguments parsed,
        IOfficeOutputWorkflowRunner runner,
        TextWriter output,
        TextWriter error,
        CancellationToken cancellationToken) {
        string inputPath = Path.GetFullPath(parsed.Inputs[0]);
        PdfPageImageExportResult result = await runner.ExportPdfPagesAsync(new PdfPageImageExportRequest {
            InputPath = inputPath,
            OutputDirectory = Path.GetFullPath(parsed.OutputPath!),
            Pages = parsed.Pages,
            Format = parsed.ImageFormat,
            TargetDpi = parsed.TargetDpi,
            MaximumDimension = parsed.MaximumDimension,
            ConflictPolicy = parsed.Force ? OfficePublicationPolicy.Replace : OfficePublicationPolicy.Fail
        }, cancellationToken: cancellationToken).ConfigureAwait(false);
        await WriteDiagnosticsAsync(result.Diagnostics, error).ConfigureAwait(false);
        await WorkflowOutput.WriteExportPagesAsync(output, result, inputPath, parsed.Json).ConfigureAwait(false);
        return result.Succeeded ? (int)OfficeImoToolExitCode.Success : MapStatus(result.Status, result.FailureKind);
    }

    private static async Task<int> AssembleAsync(
        WorkflowArguments parsed,
        IOfficeOutputWorkflowRunner runner,
        TextWriter output,
        TextWriter error,
        CancellationToken cancellationToken) {
        PdfAssemblyResult result = await runner.AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = parsed.Inputs.Select(Path.GetFullPath).ToArray(),
            OutputPath = Path.GetFullPath(parsed.OutputPath!),
            ConflictPolicy = parsed.Force ? OfficePublicationPolicy.Replace : OfficePublicationPolicy.Fail,
            Options = new PdfAssemblyOptions { IncludeSubdirectories = parsed.IncludeSubdirectories }
        }, cancellationToken: cancellationToken).ConfigureAwait(false);
        await WriteDiagnosticsAsync(result.Diagnostics, error).ConfigureAwait(false);
        await WorkflowOutput.WriteAssembleAsync(output, result, parsed.Json).ConfigureAwait(false);
        return result.Succeeded ? (int)OfficeImoToolExitCode.Success : MapStatus(result.Status, result.FailureKind);
    }

    private static async Task<int> PrintPlanAsync(
        WorkflowArguments parsed,
        TextWriter output,
        Func<PdfPrintPlanRequest, CancellationToken, Task<PdfPrintPlan>> printPlanner,
        CancellationToken cancellationToken) {
        PdfPrintPlanRequest request = parsed.CreatePrintRequest();
        PdfPrintPlan plan = await printPlanner(request, cancellationToken).ConfigureAwait(false);
        await WorkflowOutput.WritePrintPlanAsync(output, plan, request.InputPath, parsed.Json).ConfigureAwait(false);
        return (int)OfficeImoToolExitCode.Success;
    }

    private static async Task WriteDiagnosticsAsync(IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics, TextWriter error) {
        foreach (OfficeWorkflowDiagnostic diagnostic in diagnostics) {
            await error.WriteLineAsync(diagnostic.Severity + " " + diagnostic.Code + ": " + diagnostic.Message).ConfigureAwait(false);
        }
    }

    internal static int MapStatus(OfficeWorkflowStatus status, OfficeWorkflowFailureKind failureKind) => status switch {
        OfficeWorkflowStatus.Cancelled => (int)OfficeImoToolExitCode.Cancelled,
        OfficeWorkflowStatus.Failed => failureKind switch {
            OfficeWorkflowFailureKind.ValidationFailed => (int)OfficeImoToolExitCode.Usage,
            OfficeWorkflowFailureKind.InputNotFound => (int)OfficeImoToolExitCode.InputNotFound,
            OfficeWorkflowFailureKind.UnsupportedInput => (int)OfficeImoToolExitCode.UnsupportedInput,
            OfficeWorkflowFailureKind.OutputFailed => (int)OfficeImoToolExitCode.OutputFailed,
            _ => (int)OfficeImoToolExitCode.OperationFailed
        },
        _ => (int)OfficeImoToolExitCode.OperationFailed
    };
}

using System.Globalization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

internal static class WorkflowCommand {
    internal const string Usage = """
OfficeIMO.Tool - output and intake workflows

Usage:
  officeimo workflow export-pages <input.pdf> --output <folder> [--pages <selection>]
             [--format png|jpeg|webp|tiff|svg] [--dpi <36-600>] [--max-dimension <pixels>] [--force]
  officeimo workflow assemble <source>... --output <output.pdf> [--no-recursive] [--force]
  officeimo workflow print-plan <input.pdf> [--pages <selection>] [--paper A4|Letter|Legal|A3]
             [--orientation auto|portrait|landscape] [--pages-per-sheet 1|2|4]
             [--scale fit|actual|fill] [--margin <points>]

Page selections accept document-relative expressions such as 1-3,last.
Folders and ZIP archives are expanded deterministically by the reusable workflow owner.
Existing output is refused unless --force is supplied.
""";

    internal static async Task<int> RunAsync(
        string[] args,
        TextWriter standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default,
        IOfficeOutputWorkflowRunner? runner = null,
        Func<PdfPrintPlanRequest, CancellationToken, Task<PdfPrintPlan>>? printPlanner = null) {
        WorkflowCommandKind activeCommand = WorkflowCommandKind.Help;
        try {
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
        PdfPageImageExportResult result = await runner.ExportPdfPagesAsync(new PdfPageImageExportRequest {
            InputPath = Path.GetFullPath(parsed.Inputs[0]),
            OutputDirectory = Path.GetFullPath(parsed.OutputPath!),
            Pages = parsed.Pages,
            Format = parsed.ImageFormat,
            TargetDpi = parsed.TargetDpi,
            MaximumDimension = parsed.MaximumDimension,
            ConflictPolicy = parsed.Force ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail
        }, cancellationToken: cancellationToken).ConfigureAwait(false);
        await WriteDiagnosticsAsync(result.Diagnostics, error).ConfigureAwait(false);
        if (!result.Succeeded) return MapStatus(result.Status, result.FailureKind);
        await output.WriteLineAsync(result.OutputDirectory).ConfigureAwait(false);
        await output.WriteLineAsync(result.Files.Count.ToString(CultureInfo.InvariantCulture) + " page image(s), " +
                                    result.OutputBytes.ToString(CultureInfo.InvariantCulture) + " bytes").ConfigureAwait(false);
        return (int)OfficeImoToolExitCode.Success;
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
            ConflictPolicy = parsed.Force ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail,
            Options = new PdfAssemblyOptions { IncludeSubdirectories = parsed.IncludeSubdirectories }
        }, cancellationToken: cancellationToken).ConfigureAwait(false);
        await WriteDiagnosticsAsync(result.Diagnostics, error).ConfigureAwait(false);
        if (!result.Succeeded) return MapStatus(result.Status, result.FailureKind);
        await output.WriteLineAsync(result.OutputPath).ConfigureAwait(false);
        await output.WriteLineAsync(result.PageCount.ToString(CultureInfo.InvariantCulture) + " page(s) from " +
                                    result.SourceCount.ToString(CultureInfo.InvariantCulture) + " normalized source(s)").ConfigureAwait(false);
        return (int)OfficeImoToolExitCode.Success;
    }

    private static async Task<int> PrintPlanAsync(
        WorkflowArguments parsed,
        TextWriter output,
        Func<PdfPrintPlanRequest, CancellationToken, Task<PdfPrintPlan>> printPlanner,
        CancellationToken cancellationToken) {
        PdfPrintPlan plan = await printPlanner(new PdfPrintPlanRequest {
            InputPath = Path.GetFullPath(parsed.Inputs[0]),
            Pages = parsed.Pages,
            PaperSize = parsed.PaperSize,
            Orientation = parsed.Orientation,
            PagesPerSheet = parsed.PagesPerSheet,
            ScaleMode = parsed.ScaleMode,
            Margin = parsed.Margin
        }, cancellationToken).ConfigureAwait(false);
        await output.WriteLineAsync("Source pages: " + plan.SourcePageCount.ToString(CultureInfo.InvariantCulture)).ConfigureAwait(false);
        await output.WriteLineAsync("Selected pages: " + string.Join(',', plan.SelectedPages)).ConfigureAwait(false);
        await output.WriteLineAsync("Sheets: " + plan.Sheets.Count.ToString(CultureInfo.InvariantCulture)).ConfigureAwait(false);
        foreach (PdfPrintSheet sheet in plan.Sheets) {
            await output.WriteLineAsync(
                "Sheet " + sheet.SheetNumber.ToString(CultureInfo.InvariantCulture) + ": " +
                sheet.PaperSize.Width.ToString("0.##", CultureInfo.InvariantCulture) + " x " +
                sheet.PaperSize.Height.ToString("0.##", CultureInfo.InvariantCulture) + " pt; pages " +
                string.Join(',', sheet.Placements.Select(static placement => placement.PageNumber))).ConfigureAwait(false);
        }
        return (int)OfficeImoToolExitCode.Success;
    }

    private static async Task WriteDiagnosticsAsync(IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics, TextWriter error) {
        foreach (OfficeWorkflowDiagnostic diagnostic in diagnostics) {
            await error.WriteLineAsync(diagnostic.Severity + " " + diagnostic.Code + ": " + diagnostic.Message).ConfigureAwait(false);
        }
    }

    private static int MapStatus(OfficeWorkflowStatus status, OfficeWorkflowFailureKind failureKind) => status switch {
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

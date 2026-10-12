using System.Text;
using OfficeIMO.DjVu.Pdf;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odg.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Convert;

/// <summary>Maps fixed-layout command options to the shared conversion and atomic publication workflow.</summary>
internal static class FixedLayoutPdfCommand {
    internal static async Task<int> RunAsync(
        OfficePdfArguments arguments,
        string inputPath,
        string outputPath,
        Stream output,
        TextWriter error,
        CancellationToken token) {
        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = inputPath,
            OutputPath = outputPath,
            ConversionRouteId = arguments.IsDjVuInput ? "djvu-pdf" : "odg-pdf",
            ConflictPolicy = arguments.Force ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail,
            Limits = new OfficeWorkflowLimits {
                MaximumInputBytes = arguments.MaxInputBytes,
                MaximumOutputBytes = arguments.MaxOutputBytes,
                MaximumXmlCharactersInPart = arguments.MaxCharactersInPart
            },
            ConversionOptions = new OfficeWorkflowConversionOptions {
                RequireNoLoss = arguments.RequireNoLoss,
                DjVu = arguments.IsDjVuInput ? new DjVuToPdfOptions() : null,
                Draw = arguments.IsDrawInput ? new OdgToPdfOptions {
                    ForPrint = arguments.DiagramForPrint,
                    LossPolicy = OdfConversionLossPolicy.ReportOnly
                } : null
            }
        }, cancellationToken: token).ConfigureAwait(false);

        foreach (OfficeWorkflowDiagnostic diagnostic in result.Diagnostics) {
            if (diagnostic.Stage != "convert" && diagnostic.Severity == OfficeWorkflowDiagnosticSeverity.Information) continue;
            diagnostic.Details.TryGetValue("source", out string? source);
            diagnostic.Details.TryGetValue("location", out string? location);
            string context = source is null ? string.Empty : " [" + source + "]";
            if (location is not null) context += " [" + location + "]";
            await error.WriteLineAsync(diagnostic.Severity + " " + diagnostic.Code + context + ": " + diagnostic.Message).ConfigureAwait(false);
        }

        if (result.Status == OfficeWorkflowStatus.Cancelled) {
            await error.WriteLineAsync("Conversion cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        }
        if (result.Succeeded) {
            byte[] path = Encoding.UTF8.GetBytes(result.OutputPath + Environment.NewLine);
            await output.WriteAsync(path.AsMemory(), CancellationToken.None).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }

        return result.FailureKind switch {
            OfficeWorkflowFailureKind.ValidationFailed => (int)OfficeImoToolExitCode.Usage,
            OfficeWorkflowFailureKind.InputNotFound => (int)OfficeImoToolExitCode.InputNotFound,
            OfficeWorkflowFailureKind.UnsupportedInput => (int)OfficeImoToolExitCode.UnsupportedInput,
            OfficeWorkflowFailureKind.OutputFailed => (int)OfficeImoToolExitCode.OutputFailed,
            _ => (int)OfficeImoToolExitCode.OperationFailed
        };
    }
}

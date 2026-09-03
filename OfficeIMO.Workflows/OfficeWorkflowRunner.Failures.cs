namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    internal static OfficeWorkflowFailureKind ClassifyFailure(Exception exception, WorkflowFailureStage stage) {
        if (stage == WorkflowFailureStage.Output ||
            OfficeIMO.Provenance.OfficeProvenanceLimitException.IsOutput(exception) ||
            OfficeIMO.Pdf.PdfOutputLimitErrors.IsOutputLimitExceeded(exception)) {
            return OfficeWorkflowFailureKind.OutputFailed;
        }
        if (exception is FileNotFoundException or DirectoryNotFoundException) {
            return OfficeWorkflowFailureKind.InputNotFound;
        }
        if (stage == WorkflowFailureStage.Input && exception is IOException or UnauthorizedAccessException) {
            return OfficeWorkflowFailureKind.UnsupportedInput;
        }
        if (exception is NotSupportedException or InvalidDataException) {
            return OfficeWorkflowFailureKind.UnsupportedInput;
        }
        if (stage == WorkflowFailureStage.Validation || exception is ArgumentException) {
            return OfficeWorkflowFailureKind.ValidationFailed;
        }
        return OfficeWorkflowFailureKind.OperationFailed;
    }

    internal static string GetDiagnosticStage(WorkflowFailureStage stage) => stage switch {
        WorkflowFailureStage.Validation => "validate",
        WorkflowFailureStage.Input => "input",
        WorkflowFailureStage.Snapshot => "snapshot",
        WorkflowFailureStage.Output => "output",
        WorkflowFailureStage.Operation => "execute",
        _ => "execute"
    };

    internal enum WorkflowFailureStage {
        Validation,
        Input,
        Snapshot,
        Output,
        Operation
    }
}

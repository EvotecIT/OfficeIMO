namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    internal static OfficeWorkflowFailureKind ClassifyFailure(Exception exception, WorkflowFailureStage stage) {
        if (stage == WorkflowFailureStage.Output &&
            exception is IOException or UnauthorizedAccessException) {
            return OfficeWorkflowFailureKind.OutputFailed;
        }
        if (exception is FileNotFoundException or DirectoryNotFoundException) {
            return OfficeWorkflowFailureKind.InputNotFound;
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
        WorkflowFailureStage.Output => "output",
        WorkflowFailureStage.Operation => "execute",
        _ => "execute"
    };

    internal enum WorkflowFailureStage {
        Validation,
        Input,
        Output,
        Operation
    }
}

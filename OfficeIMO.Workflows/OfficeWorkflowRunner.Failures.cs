namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static OfficeWorkflowFailureKind ClassifyFailure(Exception exception, WorkflowFailureStage stage) {
        if (exception is FileNotFoundException or DirectoryNotFoundException) {
            return OfficeWorkflowFailureKind.InputNotFound;
        }
        if (exception is NotSupportedException or InvalidDataException) {
            return OfficeWorkflowFailureKind.UnsupportedInput;
        }
        if (stage == WorkflowFailureStage.Validation || exception is ArgumentException) {
            return OfficeWorkflowFailureKind.ValidationFailed;
        }
        if (stage == WorkflowFailureStage.Output &&
            exception is IOException or UnauthorizedAccessException) {
            return OfficeWorkflowFailureKind.OutputFailed;
        }
        return OfficeWorkflowFailureKind.OperationFailed;
    }

    private enum WorkflowFailureStage {
        Validation,
        Input,
        Output,
        Operation
    }
}

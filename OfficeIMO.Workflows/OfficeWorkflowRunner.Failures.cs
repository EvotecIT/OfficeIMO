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

    internal enum WorkflowFailureStage {
        Validation,
        Input,
        Output,
        Operation
    }
}

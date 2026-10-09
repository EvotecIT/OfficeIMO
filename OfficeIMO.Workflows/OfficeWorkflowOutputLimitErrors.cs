namespace OfficeIMO.Workflows;

/// <summary>Identifies output-budget failures without interpreting human-readable messages.</summary>
internal static class OfficeWorkflowOutputLimitErrors {
    private static readonly object Marker = new();

    internal static InvalidOperationException Create(string message) {
        var exception = new InvalidOperationException(message);
        exception.Data[Marker] = true;
        return exception;
    }

    internal static bool IsOutputLimitExceeded(Exception exception) {
        for (Exception? current = exception; current != null; current = current.InnerException) {
            if (current.Data.Contains(Marker)) return true;
        }
        return false;
    }
}

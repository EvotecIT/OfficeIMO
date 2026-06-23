namespace OfficeIMO.Pdf;

/// <summary>
/// Result returned by preflight-gated PDF operations.
/// </summary>
/// <typeparam name="T">Operation value type.</typeparam>
public sealed class PdfOperationResult<T> where T : class {
    private PdfOperationResult(
        string operationName,
        PdfPreflightCapability capability,
        PdfDocumentPreflight preflight,
        T? value,
        IReadOnlyList<string> diagnostics,
        Exception? exception,
        bool? canAttemptOverride) {
        OperationName = operationName;
        Capability = capability;
        Preflight = preflight;
        Value = value;
        Diagnostics = diagnostics;
        Exception = exception;
        CanAttemptOverride = canAttemptOverride;
    }

    /// <summary>Human-readable operation name.</summary>
    public string OperationName { get; }

    /// <summary>Preflight capability required by the operation.</summary>
    public PdfPreflightCapability Capability { get; }

    /// <summary>Preflight report used before attempting the operation.</summary>
    public PdfDocumentPreflight Preflight { get; }

    /// <summary>True when preflight allowed the operation to be attempted.</summary>
    public bool CanAttempt => CanAttemptOverride ?? Preflight.Can(Capability);

    /// <summary>True when the operation completed and produced a value.</summary>
    public bool Succeeded => Value is not null && Exception is null;

    /// <summary>Operation value when the operation completed.</summary>
    public T? Value { get; }

    /// <summary>Human-readable diagnostics explaining a blocked or failed operation.</summary>
    public IReadOnlyList<string> Diagnostics { get; }

    /// <summary>Exception captured from an attempted operation, when available.</summary>
    public Exception? Exception { get; }

    private bool? CanAttemptOverride { get; }

    /// <summary>Returns the operation value or throws with diagnostics when the operation failed.</summary>
    public T RequireValue() {
        if (Value is not null) {
            return Value;
        }

        string message = Diagnostics.Count == 0
            ? OperationName + " did not produce a value."
            : OperationName + " did not produce a value. " + string.Join(" ", Diagnostics);
        throw new InvalidOperationException(message, Exception);
    }

    internal static PdfOperationResult<T> Success(string operationName, PdfPreflightCapability capability, PdfDocumentPreflight preflight, T value) {
        return Success(operationName, capability, preflight, value, canAttemptOverride: null);
    }

    internal static PdfOperationResult<T> Success(string operationName, PdfPreflightCapability capability, PdfDocumentPreflight preflight, T value, bool? canAttemptOverride) {
        return new PdfOperationResult<T>(operationName, capability, preflight, value, Array.Empty<string>(), null, canAttemptOverride);
    }

    internal static PdfOperationResult<T> Blocked(string operationName, PdfPreflightCapability capability, PdfDocumentPreflight preflight) {
        return new PdfOperationResult<T>(operationName, capability, preflight, null, preflight.GetCapabilityDiagnostics(capability), null, canAttemptOverride: null);
    }

    internal static PdfOperationResult<T> Failed(string operationName, PdfPreflightCapability capability, PdfDocumentPreflight preflight, Exception exception) {
        return Failed(operationName, capability, preflight, exception, canAttemptOverride: null);
    }

    internal static PdfOperationResult<T> Failed(string operationName, PdfPreflightCapability capability, PdfDocumentPreflight preflight, Exception exception, bool? canAttemptOverride) {
        var diagnostics = new List<string>();
        AddDistinct(diagnostics, exception.Message);

        IReadOnlyList<string> capabilityDiagnostics = preflight.GetCapabilityDiagnostics(capability);
        for (int i = 0; i < capabilityDiagnostics.Count; i++) {
            AddDistinct(diagnostics, capabilityDiagnostics[i]);
        }

        return new PdfOperationResult<T>(operationName, capability, preflight, null, diagnostics.AsReadOnly(), exception, canAttemptOverride);
    }

    private static void AddDistinct(List<string> diagnostics, string? diagnostic) {
        if (string.IsNullOrWhiteSpace(diagnostic)) {
            return;
        }

        for (int i = 0; i < diagnostics.Count; i++) {
            if (string.Equals(diagnostics[i], diagnostic, StringComparison.Ordinal)) {
                return;
            }
        }

        diagnostics.Add(diagnostic!);
    }
}

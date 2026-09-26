namespace OfficeIMO.IWork;

/// <summary>One bounded-reader or semantic-projection diagnostic.</summary>
public sealed class IWorkDiagnostic {
    internal IWorkDiagnostic(IWorkDiagnosticSeverity severity, string code, string message,
        string? entryPath = null, ulong? recordIdentifier = null,
        global::OfficeIMO.OfficeConversionLossKind? lossKind = null) {
        Severity = severity;
        Code = code;
        Message = message;
        EntryPath = entryPath;
        RecordIdentifier = recordIdentifier;
        LossKind = lossKind ?? ClassifyLoss(severity, code);
    }

    /// <summary>Gets the diagnostic severity.</summary>
    public IWorkDiagnosticSeverity Severity { get; }
    /// <summary>Gets the stable diagnostic code.</summary>
    public string Code { get; }
    /// <summary>Gets the human-readable diagnostic message.</summary>
    public string Message { get; }
    /// <summary>Gets the package entry associated with the diagnostic, when available.</summary>
    public string? EntryPath { get; }
    /// <summary>Gets the IWA object identifier associated with the diagnostic, when available.</summary>
    public ulong? RecordIdentifier { get; }
    internal global::OfficeIMO.OfficeConversionLossKind LossKind { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Severity} {Code}: {Message}";

    private static global::OfficeIMO.OfficeConversionLossKind ClassifyLoss(
        IWorkDiagnosticSeverity severity, string code) {
        if (severity == IWorkDiagnosticSeverity.Information) {
            return global::OfficeIMO.OfficeConversionLossKind.None;
        }
        if (severity == IWorkDiagnosticSeverity.Error) {
            return global::OfficeIMO.OfficeConversionLossKind.Failure;
        }

        // A partially reconstructed formula still retains its typed cached value. Every
        // other current warning means source content or metadata was not represented.
        return string.Equals(code, "IWORK_TABLE_FORMULA_PARTIAL", StringComparison.Ordinal)
            ? global::OfficeIMO.OfficeConversionLossKind.Approximation
            : global::OfficeIMO.OfficeConversionLossKind.Omission;
    }
}

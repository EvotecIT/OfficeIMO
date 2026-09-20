namespace OfficeIMO.Email.Store;

/// <summary>Projects Store-specific diagnostics into the common fidelity contract.</summary>
internal static class EmailStoreFidelityProjection {
    internal static IReadOnlyList<OfficeConversionFidelityDiagnostic> Project(
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) =>
        Array.AsReadOnly(diagnostics.Select(static diagnostic => new OfficeConversionFidelityDiagnostic(
            diagnostic.Code,
            string.IsNullOrWhiteSpace(diagnostic.Message)
                ? diagnostic.Code + " was reported without a diagnostic message."
                : diagnostic.Message,
            ResolveLossKind(diagnostic),
            "OfficeIMO.Email.Store",
            diagnostic.Location)).ToArray());

    internal static OfficeConversionFidelityDiagnostic Create(
        string code,
        string message,
        OfficeConversionLossKind lossKind,
        string? location = null) =>
        new OfficeConversionFidelityDiagnostic(code, message, lossKind, "OfficeIMO.Email.Store", location);

    internal static IReadOnlyList<OfficeConversionFidelityDiagnostic> Append(
        IReadOnlyList<OfficeConversionFidelityDiagnostic> diagnostics,
        params OfficeConversionFidelityDiagnostic[] additional) =>
        additional.Length == 0
            ? diagnostics
            : Array.AsReadOnly(diagnostics.Concat(additional).ToArray());

    internal static void RequireNoLoss(IReadOnlyList<OfficeConversionFidelityDiagnostic> diagnostics) {
        OfficeConversionFidelityDiagnostic? firstLoss = diagnostics.FirstOrDefault(static diagnostic =>
            diagnostic.LossKind != OfficeConversionLossKind.None);
        if (firstLoss != null) {
            throw new InvalidDataException(
                "Email Store operation reported possible content loss. First diagnostic: " + firstLoss.Code + ".");
        }
    }

    private static OfficeConversionLossKind ResolveLossKind(EmailStoreDiagnostic diagnostic) {
        if (diagnostic.Severity == EmailStoreDiagnosticSeverity.Error ||
            diagnostic.Disposition == EmailDiagnosticDisposition.Stopped) {
            return OfficeConversionLossKind.Failure;
        }
        if (diagnostic.Disposition == EmailDiagnosticDisposition.Skipped ||
            diagnostic.DataLossRisk == EmailDataLossRisk.Confirmed) {
            return OfficeConversionLossKind.Omission;
        }
        if (diagnostic.Severity == EmailStoreDiagnosticSeverity.Warning ||
            diagnostic.Disposition == EmailDiagnosticDisposition.Recovered ||
            diagnostic.DataLossRisk == EmailDataLossRisk.Possible ||
            diagnostic.DataLossRisk == EmailDataLossRisk.Unknown) {
            return OfficeConversionLossKind.Approximation;
        }
        return OfficeConversionLossKind.None;
    }
}

namespace OfficeIMO.Email;

/// <summary>Describes known fidelity implications before an email artifact is serialized.</summary>
public sealed class EmailConversionReport : IOfficeConversionReport {
    private readonly IReadOnlyList<OfficeConversionFidelityDiagnostic> _fidelityDiagnostics;

    internal EmailConversionReport(EmailFileFormat sourceFormat, EmailFileFormat targetFormat,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        SourceFormat = sourceFormat;
        TargetFormat = targetFormat;
        Diagnostics = Array.AsReadOnly((diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).ToArray());
        _fidelityDiagnostics = Array.AsReadOnly(Diagnostics.Select(static diagnostic =>
            new OfficeConversionFidelityDiagnostic(
                diagnostic.Code,
                string.IsNullOrWhiteSpace(diagnostic.Message)
                    ? diagnostic.Code + " was reported without a diagnostic message."
                    : diagnostic.Message,
                diagnostic.LossKind,
                "OfficeIMO.Email",
                diagnostic.Location)).ToArray());
    }

    /// <summary>Format from which the in-memory document was read or created.</summary>
    public EmailFileFormat SourceFormat { get; }

    /// <summary>Requested output format.</summary>
    public EmailFileFormat TargetFormat { get; }

    /// <summary>Known fidelity and safety diagnostics for the requested conversion.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>True when the conversion is known to normalize or omit source semantics.</summary>
    public bool HasPotentialDataLoss => HasLoss;

    /// <summary>Category-preserving diagnostics for the requested format conversion.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;

    /// <summary>True when the conversion approximates, omits, or fails to preserve source content.</summary>
    public bool HasLoss => FidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>True when the active conversion policy permits serialization.</summary>
    public bool CanWrite => !Diagnostics.Any(diagnostic =>
        diagnostic.Severity == EmailDiagnosticSeverity.Error || diagnostic.Disposition == EmailDiagnosticDisposition.Stopped);

    /// <summary>Throws when the requested conversion reports possible content loss.</summary>
    public void RequireNoLoss() {
        OfficeConversionFidelityDiagnostic? firstLoss = FidelityDiagnostics.FirstOrDefault(static diagnostic =>
            diagnostic.LossKind != OfficeConversionLossKind.None);
        if (firstLoss != null) {
            throw new InvalidOperationException(
                "Email conversion reported possible content loss. First diagnostic: " + firstLoss.Code + ".");
        }
    }
}

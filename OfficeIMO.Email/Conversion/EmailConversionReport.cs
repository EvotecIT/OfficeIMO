namespace OfficeIMO.Email;

/// <summary>Describes known fidelity implications before an email artifact is serialized.</summary>
public sealed class EmailConversionReport : OfficeConversionReport {
    private readonly IReadOnlyList<OfficeDiagnostic> _fidelityDiagnostics;

    internal EmailConversionReport(EmailFileFormat sourceFormat, EmailFileFormat targetFormat,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        SourceFormat = sourceFormat;
        TargetFormat = targetFormat;
        NativeDiagnostics = Array.AsReadOnly((diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).ToArray());
        _fidelityDiagnostics = Array.AsReadOnly(NativeDiagnostics.Select(static diagnostic =>
            new global::OfficeIMO.OfficeDiagnostic(
                diagnostic.Code,
                diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.None ? global::OfficeIMO.OfficeDiagnosticSeverity.Information : global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                string.IsNullOrWhiteSpace(diagnostic.Message)
                    ? diagnostic.Code + " was reported without a diagnostic message."
                    : diagnostic.Message,
                diagnostic.Location,
                "OfficeIMO.Email",
                diagnostic.LossKind)).ToArray());
    }

    /// <summary>Format from which the in-memory document was read or created.</summary>
    public EmailFileFormat SourceFormat { get; }

    /// <summary>Requested output format.</summary>
    public EmailFileFormat TargetFormat { get; }

    /// <summary>Known fidelity and safety diagnostics for the requested conversion.</summary>
    public IReadOnlyList<EmailDiagnostic> NativeDiagnostics { get; }

    /// <summary>True when conversion fidelity includes normalization, omission, failure, or unassessed semantics.</summary>
    public bool HasPotentialDataLoss => HasLoss;

    /// <summary>Category-preserving diagnostics for the requested format conversion.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;

    /// <summary>True when the conversion approximates, omits, fails to preserve, or has not assessed source content.</summary>
    public override bool HasLoss => FidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>True when the active conversion policy permits serialization.</summary>
    public bool CanWrite => !NativeDiagnostics.Any(diagnostic =>
        diagnostic.Severity == EmailDiagnosticSeverity.Error || diagnostic.Disposition == EmailDiagnosticDisposition.Stopped);

}

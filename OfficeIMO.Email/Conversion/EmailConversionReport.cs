namespace OfficeIMO.Email;

/// <summary>Describes known fidelity implications before an email artifact is serialized.</summary>
public sealed class EmailConversionReport {
    internal EmailConversionReport(EmailFileFormat sourceFormat, EmailFileFormat targetFormat,
        IReadOnlyList<EmailDiagnostic> diagnostics, bool hasPotentialDataLoss) {
        SourceFormat = sourceFormat;
        TargetFormat = targetFormat;
        Diagnostics = diagnostics;
        HasPotentialDataLoss = hasPotentialDataLoss;
    }

    /// <summary>Format from which the in-memory document was read or created.</summary>
    public EmailFileFormat SourceFormat { get; }

    /// <summary>Requested output format.</summary>
    public EmailFileFormat TargetFormat { get; }

    /// <summary>Known fidelity and safety diagnostics for the requested conversion.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>True when the conversion is known to normalize or omit source semantics.</summary>
    public bool HasPotentialDataLoss { get; }

    /// <summary>True when the active conversion policy permits serialization.</summary>
    public bool CanWrite => !Diagnostics.Any(diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
}

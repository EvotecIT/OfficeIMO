namespace OfficeIMO.Email;

/// <summary>Result of email artifact serialization.</summary>
public sealed class EmailWriteResult {
    internal EmailWriteResult(long bytesWritten, EmailFileFormat sourceFormat, EmailFileFormat targetFormat,
        IReadOnlyList<EmailDiagnostic> diagnostics, EmailArtifactSourceSelection sourceSelection,
        EmailConversionLossDisposition lossDisposition, EmailAttachmentContentLifetime attachmentContentLifetime) {
        BytesWritten = bytesWritten;
        SourceFormat = sourceFormat;
        TargetFormat = targetFormat;
        Diagnostics = diagnostics;
        SourceSelection = sourceSelection;
        LossDisposition = lossDisposition;
        AttachmentContentLifetime = attachmentContentLifetime;
    }

    /// <summary>Number of bytes written.</summary>
    public long BytesWritten { get; }

    /// <summary>Format represented by the input model before this operation.</summary>
    public EmailFileFormat SourceFormat { get; }

    /// <summary>Requested artifact format.</summary>
    public EmailFileFormat TargetFormat { get; }

    /// <summary>Structured fidelity diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>True when serialization produced at least one error diagnostic.</summary>
    public bool HasErrors {
        get {
            foreach (EmailDiagnostic diagnostic in Diagnostics) {
                if (diagnostic.Severity == EmailDiagnosticSeverity.Error) return true;
            }
            return false;
        }
    }

    /// <summary>Whether preserved bytes or regenerated model content produced the artifact.</summary>
    public EmailArtifactSourceSelection SourceSelection { get; }

    /// <summary>Final disposition of known semantic loss.</summary>
    public EmailConversionLossDisposition LossDisposition { get; }

    /// <summary>Lifetime applied to attachment content sources during this operation.</summary>
    public EmailAttachmentContentLifetime AttachmentContentLifetime { get; }

    /// <summary>Stable diagnostic codes in emission order.</summary>
    public IReadOnlyList<string> DiagnosticCodes => Diagnostics.Select(diagnostic => diagnostic.Code).ToArray();

    /// <summary>True when the original preserved bytes were emitted verbatim.</summary>
    public bool UsedPreservedSource => SourceSelection == EmailArtifactSourceSelection.PreservedSource;

    /// <summary>Throws when the operation was blocked or accepted known semantic loss.</summary>
    public EmailWriteResult RequireNoLoss() {
        if (LossDisposition == EmailConversionLossDisposition.None && !HasErrors) return this;
        EmailDiagnostic? diagnostic = Diagnostics.FirstOrDefault(item =>
            item.Severity == EmailDiagnosticSeverity.Error || item.DataLossRisk != EmailDataLossRisk.None);
        throw new InvalidDataException(diagnostic == null
            ? "The email artifact was not produced without semantic loss."
            : string.Concat("The email artifact was not produced without semantic loss: ",
                diagnostic.Code, ": ", diagnostic.Message));
    }
}

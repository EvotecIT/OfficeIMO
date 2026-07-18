namespace OfficeIMO.Pdf;

/// <summary>
/// Immutable exact-PDF snapshot and matching internal readiness evidence for external compliance validation.
/// </summary>
public sealed class PdfComplianceArtifact {
    private readonly byte[] _bytes;
    private readonly PdfReadOptions? _readOptions;

    internal PdfComplianceArtifact(
        byte[] bytes,
        PdfComplianceReadinessReport readiness,
        PdfReadOptions? readOptions) {
        Guard.NotNull(bytes, nameof(bytes));
        Guard.NotNull(readiness, nameof(readiness));

        _bytes = bytes;
        _readOptions = readOptions;
        Readiness = readiness;
        ArtifactSha256 = PdfArtifactFingerprint.ComputeSha256(bytes);
    }

    /// <summary>Requested compliance profile captured with this exact artifact.</summary>
    public PdfComplianceProfile Profile => Readiness.Profile;

    /// <summary>Internal generation/readback readiness captured for this exact artifact.</summary>
    public PdfComplianceReadinessReport Readiness { get; }

    /// <summary>Lowercase SHA-256 of the exact artifact bytes.</summary>
    public string ArtifactSha256 { get; }

    /// <summary>Exact artifact size in bytes.</summary>
    public long ArtifactSizeBytes => _bytes.LongLength;

    /// <summary>Returns a caller-owned copy of the exact bytes to save or pass to external validators.</summary>
    public byte[] ToBytes() => (byte[])_bytes.Clone();

    /// <summary>
    /// Combines the captured readiness with external validator results bound to this exact artifact.
    /// </summary>
    public PdfComplianceProofReport AssessProof(
        IEnumerable<PdfExternalValidationResult>? externalValidations = null) =>
        PdfComplianceAnalyzer.AssessProof(Readiness, _bytes, externalValidations, _readOptions);
}

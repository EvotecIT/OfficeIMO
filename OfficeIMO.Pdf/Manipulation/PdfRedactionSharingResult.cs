namespace OfficeIMO.Pdf;

/// <summary>Verified final PDF and evidence for a redaction followed by sanitization.</summary>
public sealed class PdfRedactionSharingResult {
    internal PdfRedactionSharingResult(PdfRedactionEvidenceReport redaction, PdfSanitizationResult sanitization,
        PdfRedactionVerificationReport verification) {
        Redaction = redaction;
        Sanitization = sanitization;
        Verification = verification;
        Summary = new PdfRedactionShareableSummary(redaction, PdfRedactionPlan.ComputeSourceSha256(sanitization.ToBytes()),
            verification, sanitization.RemovedCategoryCounts.Total);
    }

    /// <summary>Evidence from the initial redaction rewrite, before sanitization.</summary>
    public PdfRedactionEvidenceReport Redaction { get; }
    /// <summary>Sanitization policy evidence and final bytes.</summary>
    public PdfSanitizationResult Sanitization { get; }
    /// <summary>Redaction checks repeated against the final sanitized artifact.</summary>
    public PdfRedactionVerificationReport Verification { get; }
    /// <summary>Exportable evidence containing counts and fingerprints, without content or free-form descriptions.</summary>
    public PdfRedactionShareableSummary Summary { get; }
    /// <summary>Returns a defensive copy of the final, sanitized PDF bytes.</summary>
    public byte[] ToBytes() => Sanitization.ToBytes();
}

/// <summary>Content-free redaction evidence suitable for serializing alongside a shared PDF.</summary>
/// <remarks>Excludes extracted text, search strings, user labels, source paths, and diagnostic messages,
/// any of which may contain the information the user intended to remove.</remarks>
public sealed class PdfRedactionShareableSummary {
    internal PdfRedactionShareableSummary(PdfRedactionEvidenceReport evidence, string outputSha256,
        PdfRedactionVerificationReport verification, int? sanitizedItemCount) {
        SourceSha256 = ToHexFingerprint(evidence.SourceSha256);
        OutputSha256 = ToHexFingerprint(outputSha256);
        PageNumbers = Array.AsReadOnly(evidence.AffectedPageNumbers.ToArray());
        AreaCount = evidence.ReviewedPlan.Areas.Count;
        VerifiedAbsentCount = evidence.VerifiedAbsentCount;
        IsVerified = evidence.IsVerified && verification.IsVerified;
        CompleteStreamInspectionRequired = verification.CompleteStreamInspectionRequired;
        ManagedRenderingChecked = verification.ManagedRenderingChecked;
        SanitizedItemCount = sanitizedItemCount;
    }

    /// <summary>Uppercase hexadecimal SHA-256 fingerprint of the reviewed input bytes.</summary>
    public string SourceSha256 { get; }
    /// <summary>Uppercase hexadecimal SHA-256 fingerprint of the final output bytes, after sanitization when requested.</summary>
    public string OutputSha256 { get; }
    /// <summary>One-based affected pages.</summary>
    public IReadOnlyList<int> PageNumbers { get; }
    /// <summary>Number of reviewed areas.</summary>
    public int AreaCount { get; }
    /// <summary>Number of content items verified absent by the redaction rewrite.</summary>
    public int VerifiedAbsentCount { get; }
    /// <summary>True when the initial redaction and final artifact checks passed.</summary>
    public bool IsVerified { get; }
    /// <summary>Whether final verification required complete stream inspection.</summary>
    public bool CompleteStreamInspectionRequired { get; }
    /// <summary>Whether final verification exercised managed rendering.</summary>
    public bool ManagedRenderingChecked { get; }
    /// <summary>Number of policy-selected metadata, payload, action, annotation, bookmark, and layer items removed, or null when sanitization was not requested.</summary>
    public int? SanitizedItemCount { get; }

    private static string ToHexFingerprint(string base64) {
#if NET6_0_OR_GREATER
        return Convert.ToHexString(Convert.FromBase64String(base64));
#else
        return BitConverter.ToString(Convert.FromBase64String(base64)).Replace("-", string.Empty);
#endif
    }
}

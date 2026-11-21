namespace OfficeIMO.Pdf;

/// <summary>Verified final PDF and evidence for a redaction followed by hidden-data removal.</summary>
public sealed class PdfRedactionSharingResult {
    internal PdfRedactionSharingResult(PdfRedactionEvidenceReport redaction, PdfHiddenDataRemoval hiddenData,
        PdfRedactionVerificationReport verification) {
        Redaction = redaction;
        HiddenData = hiddenData.Result;
        Verification = verification;
        Summary = new PdfRedactionShareableSummary(redaction, PdfRedactionPlan.ComputeSourceSha256(hiddenData.Bytes),
            verification, hiddenData.RemovedCount);
    }

    /// <summary>Evidence from the initial redaction rewrite, before hidden-data removal.</summary>
    public PdfRedactionEvidenceReport Redaction { get; }
    /// <summary>What hidden-data removal found and did, and the final bytes.</summary>
    public OfficeHiddenDataResult HiddenData { get; }
    /// <summary>Redaction checks repeated against the final artifact.</summary>
    public PdfRedactionVerificationReport Verification { get; }
    /// <summary>Exportable evidence containing counts and fingerprints, without content or free-form descriptions.</summary>
    public PdfRedactionShareableSummary Summary { get; }
    /// <summary>Returns a copy of the final PDF bytes.</summary>
    public byte[] ToBytes() => HiddenData.ToArray();
}

/// <summary>Content-free redaction evidence suitable for serializing alongside a shared PDF.</summary>
/// <remarks>Excludes extracted text, search strings, user labels, source paths, and diagnostic messages,
/// any of which may contain the information the user intended to remove.</remarks>
public sealed class PdfRedactionShareableSummary {
    internal PdfRedactionShareableSummary(PdfRedactionEvidenceReport evidence, string outputSha256,
        PdfRedactionVerificationReport verification, int? removedHiddenDataCount) {
        SourceSha256 = ToHexFingerprint(evidence.SourceSha256);
        OutputSha256 = ToHexFingerprint(outputSha256);
        PageNumbers = Array.AsReadOnly(evidence.AffectedPageNumbers.ToArray());
        AreaCount = evidence.ReviewedPlan.Areas.Count;
        VerifiedAbsentCount = evidence.VerifiedAbsentCount;
        IsVerified = evidence.IsVerified && verification.IsVerified;
        CompleteStreamInspectionRequired = verification.CompleteStreamInspectionRequired;
        ManagedRenderingChecked = verification.ManagedRenderingChecked;
        RemovedHiddenDataCount = removedHiddenDataCount;
    }

    /// <summary>Uppercase hexadecimal SHA-256 fingerprint of the reviewed input bytes.</summary>
    public string SourceSha256 { get; }
    /// <summary>Uppercase hexadecimal SHA-256 fingerprint of the final output bytes, after hidden-data removal when requested.</summary>
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
    /// <summary>Number of hidden-data findings removed, such as metadata, attachments, scripts, comments, bookmarks and layers, or null when hidden-data removal was not requested.</summary>
    public int? RemovedHiddenDataCount { get; }

    private static string ToHexFingerprint(string base64) {
#if NET6_0_OR_GREATER
        return Convert.ToHexString(Convert.FromBase64String(base64));
#else
        return BitConverter.ToString(Convert.FromBase64String(base64)).Replace("-", string.Empty);
#endif
    }
}

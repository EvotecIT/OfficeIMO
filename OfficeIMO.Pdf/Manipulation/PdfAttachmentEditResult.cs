namespace OfficeIMO.Pdf;

/// <summary>Post-save validation for one edited attachment.</summary>
public sealed class PdfAttachmentValidation {
    internal PdfAttachmentValidation(string fileName, string checksum, bool payloadMatches, bool metadataMatches) { FileName = fileName; Checksum = checksum; PayloadMatches = payloadMatches; MetadataMatches = metadataMatches; }
    /// <summary>Attachment file name.</summary>
    public string FileName { get; }
    /// <summary>Lower-case MD5 checksum stored by the PDF embedded-file contract.</summary>
    public string Checksum { get; }
    /// <summary>True when decoded readback bytes match the requested payload.</summary>
    public bool PayloadMatches { get; }
    /// <summary>True when MIME type, description, and associated-file relationship match.</summary>
    public bool MetadataMatches { get; }
    /// <summary>True when payload and metadata readback both match.</summary>
    public bool IsValid => PayloadMatches && MetadataMatches;
}

/// <summary>Edited PDF bytes with attachment and preservation proof.</summary>
public sealed class PdfAttachmentEditResult {
    private readonly byte[] _pdf;
    internal PdfAttachmentEditResult(byte[] pdf, PdfMutationPlan plan, PdfRewritePreservationReport preservation, IReadOnlyList<PdfAttachmentValidation> validations) { _pdf = (byte[])pdf.Clone(); MutationPlan = plan; PreservationReport = preservation; Validations = validations; }
    /// <summary>Shared full-rewrite mutation plan.</summary>
    public PdfMutationPlan MutationPlan { get; }
    /// <summary>Proof that non-attachment structures survived the rewrite.</summary>
    public PdfRewritePreservationReport PreservationReport { get; }
    /// <summary>Attachment payload, checksum, and metadata validations.</summary>
    public IReadOnlyList<PdfAttachmentValidation> Validations { get; }
    /// <summary>True when every requested attachment passed readback validation.</summary>
    public bool IsValid => Validations.All(static validation => validation.IsValid);
    /// <summary>Returns a defensive copy of the edited PDF.</summary>
    public byte[] ToBytes() => (byte[])_pdf.Clone();
    /// <summary>Opens the edited artifact.</summary>
    public PdfDocument ToDocument() => PdfDocument.Open(_pdf);
}

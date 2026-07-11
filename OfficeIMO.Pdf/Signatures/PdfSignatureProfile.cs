namespace OfficeIMO.Pdf;

/// <summary>High-level intent of an externally produced PDF signature.</summary>
public enum PdfSignatureProfile {
    /// <summary>Ordinary approval signature that does not certify allowed future changes.</summary>
    Approval = 0,

    /// <summary>Certification signature referenced from catalog `/Perms /DocMDP`.</summary>
    Certification = 1,

    /// <summary>RFC 3161 document timestamp emitted as `/Type /DocTimeStamp`.</summary>
    DocumentTimestamp = 2
}

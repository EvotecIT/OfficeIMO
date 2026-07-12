namespace OfficeIMO.Pdf;

/// <summary>Standard password-security algorithms supported for PDF output.</summary>
public enum PdfStandardEncryptionAlgorithm {
    /// <summary>AES-256 with Standard security handler revision 6 and AESV3 crypt filters.</summary>
    Aes256,

    /// <summary>AES-128 with Standard security handler revision 4 and AESV2 crypt filters.</summary>
    Aes128,

    /// <summary>Legacy RC4-128 with Standard security handler revision 3.</summary>
    LegacyRc4
}

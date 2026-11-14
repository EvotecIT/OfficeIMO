namespace OfficeIMO.Pdf;

/// <summary>
/// Thrown when authenticated decryption (AES-GCM, crypt filter /AESV4) rejects an encrypted string or stream:
/// the content was modified after encryption or does not belong to this document.
/// </summary>
public sealed class PdfEncryptedContentIntegrityException : PdfEncryptionException {
    /// <summary>Creates an encrypted-content integrity exception.</summary>
    public PdfEncryptedContentIntegrityException(string message) : base(message) {
    }
}

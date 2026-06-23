namespace OfficeIMO.Pdf;

/// <summary>Base exception for encrypted PDFs that cannot be opened with the supplied read options.</summary>
public class PdfEncryptionException : NotSupportedException {
    /// <summary>Creates a PDF encryption exception.</summary>
    public PdfEncryptionException(string message) : base(message) {
    }
}

/// <summary>Thrown when an encrypted PDF requires a password and none was supplied or the empty password was not accepted.</summary>
public sealed class PdfPasswordRequiredException : PdfEncryptionException {
    /// <summary>Creates a password-required exception.</summary>
    public PdfPasswordRequiredException(string message) : base(message) {
    }
}

/// <summary>Thrown when a supplied PDF password does not authenticate.</summary>
public sealed class PdfInvalidPasswordException : PdfEncryptionException {
    /// <summary>Creates an invalid-password exception.</summary>
    public PdfInvalidPasswordException(string message) : base(message) {
    }
}

/// <summary>Thrown when a PDF uses an encryption mode that OfficeIMO.Pdf does not support yet.</summary>
public sealed class PdfUnsupportedEncryptionException : PdfEncryptionException {
    /// <summary>Creates an unsupported-encryption exception.</summary>
    public PdfUnsupportedEncryptionException(string message) : base(message) {
    }
}

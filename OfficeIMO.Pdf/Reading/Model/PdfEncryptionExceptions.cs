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

/// <summary>Thrown when a public-key encrypted PDF is opened without <see cref="PdfLoadOptions.PublicKeyDecryption"/>.</summary>
public sealed class PdfRecipientCertificateRequiredException : PdfEncryptionException {
    /// <summary>Creates a recipient-certificate-required exception.</summary>
    public PdfRecipientCertificateRequiredException(string message) : base(message) {
    }
}

/// <summary>Thrown when none of the supplied certificates is a recipient of a public-key encrypted PDF.</summary>
public sealed class PdfRecipientNotFoundException : PdfEncryptionException {
    /// <summary>Creates a recipient-not-found exception.</summary>
    public PdfRecipientNotFoundException(string message) : base(message) {
    }
}

/// <summary>
/// Thrown when a recipient entry of a public-key encrypted PDF could not be decrypted with a supplied certificate,
/// for example because its private key is missing or the envelope uses an algorithm the provider does not support.
/// </summary>
public sealed class PdfRecipientDecryptionException : PdfEncryptionException {
    /// <summary>Creates a recipient-decryption exception.</summary>
    public PdfRecipientDecryptionException(string message) : base(message) {
    }
}

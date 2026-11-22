namespace OfficeIMO.Reader.FormatInternals;

/// <summary>Maps neutral Reader admission settings to the canonical package security owner.</summary>
internal static partial class OpenXmlReaderAdmission {
    /// <summary>
    /// Reports a structurally damaged package as damaged content rather than a general I/O failure: one the Open XML SDK
    /// cannot open, or one the package security check rejects as malformed. Policy and size rejections keep their codes.
    /// </summary>
    internal static string? ClassifyReadError(Exception exception) => exception switch {
        DocumentFormat.OpenXml.Packaging.OpenXmlPackageException => ReaderErrorCodes.InvalidContent,
        OfficePackageSecurityException { Rule: OfficePackageSecurityRule.MalformedPackage or OfficePackageSecurityRule.MalformedRelationship
            or OfficePackageSecurityRule.DuplicatePartName or OfficePackageSecurityRule.UnsafePartName } => ReaderErrorCodes.InvalidContent,
        _ => null
    };

    internal static bool CanRetryEncrypted(Exception exception, ReaderOptions options) =>
        !string.IsNullOrEmpty(options.OpenPassword)
        && exception is not OfficePackageSecurityException and not OfficeInspectionResourceLimitException and not ReaderResourceLimitException
        && exception is InvalidDataException or IOException or System.Security.Cryptography.CryptographicException;
}

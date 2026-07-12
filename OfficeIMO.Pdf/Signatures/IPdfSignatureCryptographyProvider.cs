namespace OfficeIMO.Pdf;

/// <summary>Optional cryptography seam for CMS, certificate-chain, timestamp, and revocation validation.</summary>
public interface IPdfSignatureCryptographyProvider {
    /// <summary>Stable provider name included in signature reports.</summary>
    string Name { get; }

    /// <summary>Validates cryptographic evidence prepared by the dependency-free PDF parser.</summary>
    PdfSignatureCryptographicResult Verify(PdfSignatureCryptographyInput input);
}

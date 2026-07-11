namespace OfficeIMO.Pdf;

/// <summary>Caller-owned cloud, HSM, smart-card, or local signing callback.</summary>
public interface IPdfExternalSigner {
    /// <summary>Stable signer implementation name included in completion diagnostics.</summary>
    string Name { get; }

    /// <summary>Produces CMS, CAdES, or RFC 3161 bytes for the prepared PDF byte ranges.</summary>
    byte[] Sign(PdfExternalSignatureRequest request);
}

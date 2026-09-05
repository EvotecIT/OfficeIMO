using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Editor;

public sealed record PdfSigningCertificateViewModel(
    string Thumbprint,
    string DisplayName,
    DateTime NotAfter,
    string Issuer) {
    public string Label => DisplayName + " · expires " + NotAfter.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
}

public sealed record PdfBatesPositionChoice(PdfBatesPosition Position, string Label);

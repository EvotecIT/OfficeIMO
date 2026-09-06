using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Editor;

internal sealed record PdfSigningSettings(PdfSigningCertificateViewModel Certificate, string FieldName, string Reason,
    string Location, bool Visible, int Page, double X, double Y, double Width, double Height) {
    internal PdfExternalSignatureOptions CreateOptions() => new() {
        FieldName = FieldName, Name = Certificate.DisplayName,
        Reason = string.IsNullOrWhiteSpace(Reason) ? null : Reason.Trim(),
        Location = string.IsNullOrWhiteSpace(Location) ? null : Location.Trim(),
        VisibleAppearance = Visible ? new() {
            PageNumber = Page, X = X, Y = Y, Width = Width, Height = Height,
            Text = "Digitally signed by " + Certificate.DisplayName
        } : null
    };
}

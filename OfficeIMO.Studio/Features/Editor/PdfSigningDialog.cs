using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Editor;

/// <summary>Desktop presentation of the shared document review.</summary>
public sealed class PdfSigningDialog : StudioDialogWindow {
    public PdfSigningDialog() : base(new PdfSigningDialogContent()) { }
    internal PdfSigningDialog(PdfSigningPreviewViewModel model) : base(new PdfSigningDialogContent(model)) { }
}

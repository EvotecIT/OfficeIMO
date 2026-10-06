using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Editor;

/// <summary>Desktop presentation of the shared document review.</summary>
public sealed class PdfProtectionDialog : StudioDialogWindow {
    public PdfProtectionDialog() : base(new PdfProtectionDialogContent()) { }
    internal PdfProtectionDialog(PdfProtectionPreviewViewModel model) : base(new PdfProtectionDialogContent(model)) { }
}

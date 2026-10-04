using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Organizer;

/// <summary>Desktop presentation of the shared document review.</summary>
public sealed class PageImportDialog : StudioDialogWindow {
    public PageImportDialog() : base(new PageImportDialogContent()) { }
    internal PageImportDialog(PageImportPreviewViewModel model) : base(new PageImportDialogContent(model)) { }
}

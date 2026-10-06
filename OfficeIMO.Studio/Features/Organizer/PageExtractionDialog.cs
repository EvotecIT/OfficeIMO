using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Organizer;

/// <summary>Desktop presentation of the shared document review.</summary>
public sealed class PageExtractionDialog : StudioDialogWindow {
    public PageExtractionDialog() : base(new PageExtractionDialogContent()) { }
    internal PageExtractionDialog(PageExtractionPreviewViewModel model) : base(new PageExtractionDialogContent(model)) { }
}

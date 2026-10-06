using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Organizer;

/// <summary>Desktop presentation of the shared document review.</summary>
public sealed class PageMoveDialog : StudioDialogWindow {
    public PageMoveDialog() : base(new PageMoveDialogContent()) { }
    internal PageMoveDialog(PageMovePreviewViewModel model) : base(new PageMoveDialogContent(model)) { }
}

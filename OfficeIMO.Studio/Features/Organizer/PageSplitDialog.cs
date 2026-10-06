using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Organizer;

/// <summary>Desktop presentation of the shared document review.</summary>
public sealed class PageSplitDialog : StudioDialogWindow {
    public PageSplitDialog() : base(new PageSplitDialogContent()) { }
    internal PageSplitDialog(PageSplitPreviewViewModel model) : base(new PageSplitDialogContent(model)) { }
}

using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Editor;

/// <summary>Desktop presentation of the shared document review.</summary>
public sealed class WatermarkDialog : StudioDialogWindow {
    public WatermarkDialog() : base(new WatermarkDialogContent()) { }
    public WatermarkDialog(WatermarkPreviewViewModel model) : base(new WatermarkDialogContent(model)) {
        Opened += (_, _) => {
            if (Owner is not { } owner) return;
            Width = Math.Max(MinWidth, Math.Min(Width, owner.Bounds.Width - 40));
            Height = Math.Max(MinHeight, Math.Min(Height, owner.Bounds.Height - 40));
        };
    }
}

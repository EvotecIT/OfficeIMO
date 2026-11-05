using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Sets a solid background, preserving native theme references and literal transparency.</summary>
    public void SetBackgroundColor(OfficeColor color) {
        BackgroundColor = color;
    }

    /// <summary>Removes the slide's background override so the layout or master can supply its background.</summary>
    public void ClearBackground() {
        ClearBackgroundImage();
        CommonSlideData? common = SlideRoot.CommonSlideData;
        if (common != null) common.Background = null;
    }
}

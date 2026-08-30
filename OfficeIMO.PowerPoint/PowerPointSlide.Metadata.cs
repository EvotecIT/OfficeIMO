using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Gets or sets the authored slide name stored in common slide data.</summary>
    public string? Name {
        get => SlideRoot.CommonSlideData?.Name?.Value;
        set {
            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            if (string.IsNullOrEmpty(value)) data.RemoveAttribute("name", string.Empty);
            else data.Name = value;
        }
    }
}

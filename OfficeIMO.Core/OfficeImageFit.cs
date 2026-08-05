namespace OfficeIMO.Drawing;

/// <summary>Reusable image fitting modes for placing raster images inside a target box.</summary>
public enum OfficeImageFit {
    /// <summary>Scale the image to exactly fill the target width and height.</summary>
    Stretch = 0,

    /// <summary>Scale the image proportionally so the whole image fits inside the target box.</summary>
    Contain = 1,

    /// <summary>Scale the image proportionally so the target box is fully covered, clipping any overflow.</summary>
    Cover = 2
}

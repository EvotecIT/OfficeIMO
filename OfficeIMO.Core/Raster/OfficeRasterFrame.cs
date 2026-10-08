using System;

namespace OfficeIMO.Drawing;

/// <summary>A rendered raster frame or page and its display duration.</summary>
public sealed class OfficeRasterFrame {
    /// <summary>Creates a frame referring to a mutable managed image.</summary>
    /// <remarks>The buffer is retained by reference. Call <see cref="OfficeRasterImage.Clone"/> when an independent snapshot is needed.</remarks>
    public OfficeRasterFrame(OfficeRasterImage image, TimeSpan duration = default) {
        Image = image ?? throw new ArgumentNullException(nameof(image));
        if (duration < TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(duration));
        Duration = duration;
    }

    /// <summary>Rendered full-canvas image for animation, or the independent image for a page.</summary>
    public OfficeRasterImage Image { get; }
    /// <summary>Display duration. Static images and pages use zero.</summary>
    public TimeSpan Duration { get; }
}

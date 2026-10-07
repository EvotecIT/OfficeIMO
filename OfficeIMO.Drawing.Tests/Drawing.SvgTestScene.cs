using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

internal static class SvgTestScene {
    // Metadata assertions inspect the content inside the authored root viewport.
    // Rendering tests retain the complete clipped scene.
    internal static OfficeDrawing Content(OfficeDrawing drawing) {
        var viewport = Assert.IsType<OfficeDrawingGroup>(Assert.Single(drawing.Elements));
        Assert.Equal(drawing.Width, viewport.ClipPath.Width);
        Assert.Equal(drawing.Height, viewport.ClipPath.Height);
        Assert.Equal(0D, viewport.X);
        Assert.Equal(0D, viewport.Y);
        return viewport.InnerDrawing;
    }
}

using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointCustomGeometryLengthTests {
    [Fact]
    public void CustomGeometryRetainsMixedUnitsAndWholeEmuBoundsAfterReload() {
        var path = OfficeShape.Path(OfficePathCommand.MoveTo(0, 0),
            OfficePathCommand.LineTo(100, 100));
        using var bytes = new MemoryStream();
        using (var document = PowerPointDocument.Create()) {
            document.AddSlide().AddCustomGeometry(path, OfficeLength.Emus(-1), OfficeLength.Points(12.5),
                OfficeLength.Emus(914401), OfficeLength.Centimeters(2.54));
            document.Save(bytes);
        }
        bytes.Position = 0;
        using var loaded = PowerPointDocument.Load(bytes);
        var shape = Assert.IsType<PowerPointAutoShape>(Assert.Single(loaded.Slides[0].Shapes));
        Assert.Equal(-1L, shape.Left.ToEmus());
        Assert.Equal(158750L, shape.Top.ToEmus());
        Assert.Equal(914401L, shape.Width.ToEmus());
        Assert.Equal(914400L, shape.Height.ToEmus());
    }
}

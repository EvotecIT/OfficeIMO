using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingSvgGeometryBudgetTests {
    [Fact]
    public void FixedPageImportCanExplicitlyIncreaseGeometryBudgetWithoutChangingDefault() {
        string geometry = "M0 0 " + string.Concat(Enumerable.Repeat("L1 1 L2 1 L1 2 ", 7000)) + "Z";
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'><path d='" + geometry + "'/></svg>");
        Assert.False(OfficeSvgDrawingReader.TryRead(svg, out _));
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, new OfficeSvgDrawingReaderOptions { MaximumGeometryCommands = 22000 }, out var drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.False(OfficeSvgDrawingReader.TryRead(svg, new OfficeSvgDrawingReaderOptions { MaximumGeometryCommands = 21000 }, out _));
        Assert.False(OfficeSvgDrawingReader.TryRead(svg, new OfficeSvgDrawingReaderOptions { MaximumGeometryCommands = 1000001 }, out _));
    }
    [Fact]
    public void GeometryAllowanceIncludesDefinitionsAndClips() {
        string path = "M0 0 L1 0 L1 1 Z";
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'><defs><clipPath id='c'><path d='" + path + "'/></clipPath></defs><path clip-path='url(#c)' d='" + path + "'/></svg>");
        Assert.False(OfficeSvgDrawingReader.TryRead(svg, new OfficeSvgDrawingReaderOptions { MaximumGeometryCommands = 7 }, out _));
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, new OfficeSvgDrawingReaderOptions { MaximumGeometryCommands = 8 }, out _, out int unsupported));
        Assert.Equal(0, unsupported);
    }
    [Fact]
    public void ExpandedReferencesCannotExceedTheGeometryAllowance() {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'><defs><path id='p' d='M0 0 L1 0 L1 1 Z'/></defs><use href='#p'/><use href='#p'/><use href='#p'/></svg>");
        Assert.False(OfficeSvgDrawingReader.TryRead(svg, new OfficeSvgDrawingReaderOptions { MaximumGeometryCommands = 8 }, out _));
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, new OfficeSvgDrawingReaderOptions { MaximumGeometryCommands = 12 }, out _, out int unsupported));
        Assert.Equal(0, unsupported);
    }

}

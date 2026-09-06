using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgClippingTests {
    [Fact]
    public void SvgTransformedClipRetainsPaintOutsideItsUntransformedCanvas() {
        OfficeRasterImage image = ReadClippedSvg("<defs><clipPath id='c' transform='translate(-20)'><rect x='40' width='20' height='20'/></clipPath></defs>" +
            "<rect width='40' height='20' fill='red' clip-path='url(#c)'/>");
        Assert.Equal(OfficeColor.Red, image.GetPixel(30, 10));
        Assert.Equal(0, image.GetPixel(10, 10).A);
    }

    [Fact]
    public void SvgNestedViewportClipUsesItsLocalCoordinateSystem() {
        OfficeRasterImage image = ReadClippedSvg("<defs><clipPath id='c'><rect width='10' height='20'/></clipPath></defs>" +
            "<svg x='20' width='20' height='20' clip-path='url(#c)'><rect width='20' height='20' fill='red'/></svg>");
        Assert.Equal(OfficeColor.Red, image.GetPixel(25, 10));
        Assert.Equal(0, image.GetPixel(35, 10).A);
    }

    [Theory]
    [InlineData("<svg x='20' width='20' height='20' viewBox='10 5 10 10' clip-path='url(#c)'><rect x='10' y='5' width='10' height='10' fill='red'/></svg>", "<rect x='10' y='5' width='5' height='10'/>")]
    [InlineData("<defs><rect id='paint' width='20' height='20' fill='red'/></defs><use href='#paint' x='20' clip-path='url(#c)'/>", "<rect width='10' height='20'/>")]
    [InlineData("<defs><symbol id='paint' viewBox='0 0 100 100'><rect width='100' height='100' fill='red'/></symbol></defs><use href='#paint' x='20' width='20' height='20' clip-path='url(#c)'/>", "<rect width='10' height='20'/>")]
    public void SvgClipUsesViewportAndReferencePlacement(string content, string clip) {
        OfficeRasterImage image = ReadClippedSvg("<defs><clipPath id='c'>" + clip + "</clipPath></defs>" + content);
        Assert.Equal(OfficeColor.Red, image.GetPixel(25, 10));
        Assert.Equal(0, image.GetPixel(35, 10).A);
    }

    [Theory]
    [InlineData("<rect width='0' height='20'/>")]
    [InlineData("<circle r='0'/>")]
    [InlineData("<ellipse rx='5' ry='0'/>")]
    [InlineData("<path d=''/>")]
    [InlineData("<path d='M0 0'/>")]
    public void SvgEmptyClipGeometrySuppressesAllPaint(string geometry) {
        OfficeRasterImage image = ReadClippedSvg("<defs><clipPath id='c'>" + geometry + "</clipPath></defs>" +
            "<rect width='40' height='20' fill='red' clip-path='url(#c)'/>");
        Assert.Equal(0, image.GetPixel(10, 10).A);
    }

    [Fact]
    public void SvgClipImplicitlyClosesOpenPathContours() {
        OfficeRasterImage image = ReadClippedSvg("<defs><clipPath id='c'><path d='M0 0 L40 0 L40 20'/></clipPath></defs>" +
            "<rect width='40' height='20' fill='red' clip-path='url(#c)'/>");
        Assert.Equal(OfficeColor.Red, image.GetPixel(30, 5));
        Assert.Equal(0, image.GetPixel(5, 15).A);
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("group")]
    [InlineData("effect")]
    [InlineData("blend")]
    [InlineData("mask")]
    [InlineData("stroke")]
    public void SvgViewportRetainsTransformedGeometryBeyondTheViewBox(string kind) {
        string content = kind switch {
            "group" => "<g transform='translate(-10)'><rect width='20' height='20' fill='red'/></g>",
            "effect" => "<g transform='translate(-10)' opacity='.5'><rect width='20' height='20' fill='red'/></g>",
            "blend" => "<g transform='translate(-10)' style='mix-blend-mode:multiply'><rect width='20' height='20' fill='red'/></g>",
            "mask" => "<defs><mask id='m' maskUnits='userSpaceOnUse' x='-10' y='0' width='40' height='20'><rect x='-10' width='40' height='20' fill='white'/></mask></defs>" +
                "<g transform='translate(-10)' mask='url(#m)'><rect width='20' height='20' fill='red'/></g>",
            "stroke" => "<rect x='0' y='0' width='10' height='20' fill='none' stroke='red' stroke-width='20'/>",
            _ => "<rect width='20' height='20' transform='translate(-10)' fill='red'/>"
        };
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20' viewBox='0 0 20 20'>" +
            content + "</svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        SaveClipEvidence(svg, drawing!, "viewport-" + kind);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.Equal(kind == "effect" ? OfficeColor.FromRgba(255, 0, 0, 128) : OfficeColor.Red, image.GetPixel(5, 10));
        Assert.Equal(kind == "effect" ? OfficeColor.FromRgba(255, 0, 0, 128) : OfficeColor.Red, image.GetPixel(15, 10));
        if (kind != "stroke") Assert.Equal(0, image.GetPixel(25, 10).A);
    }

    [Fact]
    public void SvgOverflowExpansionRespectsTheConfiguredViewportBudget() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20' viewBox='0 0 100 100'>" +
            "<rect x='-50' width='200' height='100' fill='red'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumViewportDimension = 100, MaximumViewportPixels = 10000 };
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    private static OfficeRasterImage ReadClippedSvg(string content) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'>" + content + "</svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        SaveClipEvidence(svg, drawing!, "clip");
        return OfficeDrawingRasterRenderer.Render(drawing!);
    }

    private static void SaveClipEvidence(string svg, OfficeDrawing drawing, string name) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_SVG_CLIPPING_ARTIFACTS");
        if (string.IsNullOrWhiteSpace(output)) return;
        using var hash = System.Security.Cryptography.SHA256.Create();
        name += "-" + BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(svg))).Replace("-", "").Substring(0, 12);
        Directory.CreateDirectory(output);
        File.WriteAllText(Path.Combine(output, name + ".source.svg"), svg);
        File.WriteAllText(Path.Combine(output, name + ".drawing.svg"), OfficeDrawingSvgExporter.ToSvg(drawing));
        File.WriteAllBytes(Path.Combine(output, name + ".png"), OfficeDrawingRasterRenderer.ToPng(drawing, 1D, OfficeColor.White));
    }

    [Theory]
    [InlineData("root")]
    [InlineData("group")]
    [InlineData("shape")]
    public void SvgUserSpaceClipRestrictsPaintAtEachElementBoundary(string scope) {
        const string clip = "clip-path='url(#window)'";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='30' " + (scope == "root" ? clip : "") + ">" +
            "<defs><clipPath id='window'><rect x='10' y='5' width='20' height='20'/></clipPath></defs>" +
            "<g " + (scope == "group" ? clip : "") + "><rect width='40' height='30' fill='red' " + (scope == "shape" ? clip : "") + "/></g></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.Equal(OfficeColor.Red, image.GetPixel(20, 15));
        Assert.Equal(0, image.GetPixel(2, 15).A);
        Assert.Equal(0, image.GetPixel(20, 2).A);
    }

    [Theory]
    [InlineData("clipPathUnits='objectBoundingBox'", "<rect width='1' height='1'/>")]
    [InlineData("", "<circle cx='10' cy='10' r='5'/><circle cx='20' cy='10' r='5'/>")]
    public void SvgClipGeometryOutsideTheSupportedSubsetIsDiagnosed(string attributes, string geometry) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='30'><defs>" +
            "<clipPath id='clip' " + attributes + ">" + geometry + "</clipPath></defs>" +
            "<rect width='40' height='30' fill='red' clip-path='url(#clip)'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Theory]
    [InlineData("svg")]
    [InlineData("symbol")]
    public void SvgViewportRetainsGeometryCrossingTheViewBoxUntilTheViewportClip(string kind) {
        const string content = "<rect x='-10' y='0' width='40' height='20' fill='red'/>";
        string child = kind == "svg"
            ? "<svg x='5' y='5' width='40' height='20' viewBox='0 0 20 20'>" + content + "</svg>"
            : "<defs><symbol id='paint' viewBox='0 0 20 20'>" + content + "</symbol></defs><use href='#paint' x='5' y='5' width='40' height='20'/>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='50' height='30'>" + child + "</svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.Equal(OfficeColor.Red, image.GetPixel(7, 15));
        Assert.Equal(OfficeColor.Red, image.GetPixel(42, 15));
        Assert.Equal(0, image.GetPixel(2, 15).A);
        Assert.Equal(0, image.GetPixel(47, 15).A);
    }
}

using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeImageReaderReadsSvgWithExternalDtdWithoutResolvingIt() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<?xml version='1.0'?><!DOCTYPE svg PUBLIC '-//W3C//DTD SVG 1.1//EN' 'https://invalid.example.test/svg11.dtd'>"
            + "<svg xmlns='http://www.w3.org/2000/svg' width='120' height='100' viewBox='0 0 120 100'><rect width='120' height='100'/></svg>");

        Assert.True(OfficeImageReader.TryIdentify(svg, null, out OfficeImageInfo sniffed));
        Assert.Equal(OfficeImageFormat.Svg, sniffed.Format);
        Assert.Equal(120, sniffed.Width);
        Assert.Equal(100, sniffed.Height);
        Assert.True(OfficeImageReader.TryIdentifyByContent(svg, null, out OfficeImageInfo identified));
        Assert.Equal(120, identified.Width);
        Assert.Equal(100, identified.Height);
        Assert.True(OfficeImageReader.TryValidateContent(svg, null, out OfficeImageInfo validated));
        Assert.Equal(120, validated.Width);
        Assert.Equal(100, validated.Height);
    }

    [Fact]
    public void OfficeImageReaderKeepsHeaderProbeAfterLongSvgComment() {
        byte[] ordinary = Encoding.UTF8.GetBytes("<!--" + new string('x', 5000)
            + "--><svg xmlns='http://www.w3.org/2000/svg' width='3' height='2'><unclosed");
        Assert.True(OfficeImageReader.TryIdentify(ordinary, "header.svg", out OfficeImageInfo header));
        Assert.Equal(3, header.Width);
        Assert.Equal(2, header.Height);

        byte[] withDtd = Encoding.UTF8.GetBytes("<!--" + new string('x', 5000)
            + "--><!DOCTYPE svg [<!ENTITY secret SYSTEM 'file:///should-not-resolve'>]>"
            + "<svg xmlns='http://www.w3.org/2000/svg' width='3' height='2'>&secret;</svg>");
        Assert.True(OfficeImageReader.TryIdentify(withDtd, "header.svg", out OfficeImageInfo fallback));
        Assert.Equal(0, fallback.Width);
        Assert.Equal(0, fallback.Height);
        Assert.False(OfficeImageReader.TryValidateContent(withDtd, "header.svg", out _));
    }

}

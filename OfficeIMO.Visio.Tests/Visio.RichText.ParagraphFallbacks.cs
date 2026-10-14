using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioRichTextParagraphRegressionTests {
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 2)]
    [InlineData(true, 2)]
    public void UnsupportedParagraphPlacementRetainsCharacterFormattingAndAlignment(bool connector, int alignment) {
        VisioPage control = Synthetic(connector, alignment, "<SpBefore>0</SpBefore>").Pages[0];
        double expectedLeft = Number(Label(control).Descendants(Svg + "text").First(), "x");
        int expectedRasterLeft = AssertRaster(control, checkBackground: false);
        foreach (string cells in new[] { "<SpBefore>-0.125</SpBefore>", "<SpAfter>-0.125</SpAfter>", "<IndFirst>-0.125</IndFirst>" }) {
            VisioDocument document = Synthetic(connector, alignment, cells);
            foreach (VisioDocument candidate in Variants(document)) {
                byte[] native = candidate.ToLegacyXmlResult().Value;
                XElement[] text = Label(candidate.Pages[0]).Descendants(Svg + "text").ToArray();
                Assert.Equal(new[] { "RED", "GREEN" }, text.Select(node => node.Value));
                Assert.Equal("#CC1122", (string?)text[0].Attribute("fill"));
                Assert.Equal("#117733", (string?)text[1].Attribute("fill"));
                Assert.Equal("700", (string?)text[0].Attribute("font-weight"));
                Assert.Equal("italic", (string?)text[1].Attribute("font-style"));
                Assert.InRange(Number(text[0], "x") - expectedLeft, -.002D, .002D);
                Assert.Equal(expectedRasterLeft, AssertRaster(candidate.Pages[0], checkBackground: false));
                Assert.Equal(native, candidate.ToLegacyXmlResult().Value);
            }
        }
    }

    [Theory]
    [InlineData(false, 1, "<IndLeft>0</IndLeft>")]
    [InlineData(true, 1, "<IndLeft>0</IndLeft>")]
    [InlineData(false, 2, "<IndLeft>0</IndLeft>")]
    [InlineData(true, 2, "<IndLeft>0</IndLeft>")]
    [InlineData(false, 0, "<IndLeft>0.5</IndLeft>")]
    [InlineData(true, 0, "<IndLeft>0.5</IndLeft>")]
    public void ParagraphBackgroundSurroundsPlacedTextInsteadOfItsAlignmentInset(bool connector, int alignment, string cells) {
        foreach (VisioDocument candidate in Variants(Synthetic(connector, alignment, cells))) {
            XElement label = Label(candidate.Pages[0]);
            XElement background = Assert.Single(label.Descendants(Svg + "rect"));
            double left = label.Descendants(Svg + "text").Min(node => Number(node, "x"));
            Assert.InRange(Number(background, "x") - left, -3.002D, -2.998D);
            Assert.InRange(Number(background, "width"), 10D, 150D);
            AssertRaster(candidate.Pages[0], checkBackground: true);
        }
    }

    private static int AssertRaster(VisioPage page, bool checkBackground) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ToPng(new VisioPngSaveOptions { Supersampling = 1 }), out var image));
        int glyphLeft = image!.Width, backgroundLeft = image.Width;
        bool red = false, green = false;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor color = image.GetPixel(x, y);
                bool isRed = color.R > 100 && color.R > color.G * 2 && color.R > color.B * 2;
                bool isGreen = color.G > 50 && color.G > color.R * 2 && color.G > color.B * 1.5;
                red |= isRed; green |= isGreen;
                if (isRed || isGreen) glyphLeft = Math.Min(glyphLeft, x);
                if (color.R > 200 && color.G > 200 && color.B < 20) backgroundLeft = Math.Min(backgroundLeft, x);
            }
        }
        Assert.True(red && green, "Both native character colors must reach raster export.");
        if (checkBackground) Assert.InRange(backgroundLeft, glyphLeft - 6, glyphLeft - 1);
        return glyphLeft;
    }

    private static IEnumerable<VisioDocument> Variants(VisioDocument document) {
        yield return document;
        yield return VisioDocument.Load(new MemoryStream(document.ToBytes()));
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    }

    private static XElement Label(VisioPage page) => XDocument.Parse(page.ToSvg()).Descendants(Svg + "g")
        .Single(node => (string?)node.Attribute("data-visio-shape-id") == "1" || (string?)node.Attribute("data-visio-connector-id") == "1");

    private static double Number(XElement node, string name) => double.Parse(node.Attribute(name)!.Value, CultureInfo.InvariantCulture);

    private static VisioDocument Synthetic(bool connector, int alignment, string cells) {
        string transform = connector
            ? "<XForm1D><BeginX>1</BeginX><BeginY>2</BeginY><EndX>5</EndX><EndY>2</EndY></XForm1D><TextXForm><TxtWidth>4</TxtWidth><TxtHeight>2</TxtHeight></TextXForm>"
            : "<XForm><PinX>3</PinX><PinY>2</PinY><Width>4</Width><Height>2</Height></XForm>";
        string source = $"<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Pages><Page ID='0'><PageSheet><PageProps><PageWidth>6</PageWidth><PageHeight>4</PageHeight></PageProps></PageSheet><Shapes><Shape ID='1'>{transform}<Fill><FillPattern>0</FillPattern></Fill><Line><LinePattern>0</LinePattern></Line><TextBlock><TextBkgnd>#FFFF01</TextBkgnd><VerticalAlign>0</VerticalAlign></TextBlock><Char IX='7'><Size>0.16666666666666667</Size><Color>#cc1122</Color><Style>1</Style></Char><Char IX='9'><Size>0.16666666666666667</Size><Color>#117733</Color><Style>2</Style></Char><Para IX='0'><HorzAlign>{alignment}</HorzAlign><SpLine>-1.2</SpLine>{cells}</Para><Text><pp IX='0'/><cp IX='7'/>RED<cp IX='9'/>GREEN</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
    }
}

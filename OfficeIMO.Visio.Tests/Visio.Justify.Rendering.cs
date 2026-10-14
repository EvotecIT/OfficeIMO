using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioJustifyRenderingTests {
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeJustifiedParagraphsExpandWrappedLinesAndKeepTheFinalLineLeftAligned(bool connector) {
        VisioPage left = NativeParagraph(connector, alignment: 0).Pages[0];
        VisioPage justified = NativeParagraph(connector, alignment: 3).Pages[0];
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A', 'B', 'C', 'D', 'E', 'F', 'G');
        var svgOptions = new VisioSvgSaveOptions { ResolveConnectorLabelOverlaps = false };
        svgOptions.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, font);
        XElement[][] leftLines = SvgLines(left, svgOptions);
        XElement[][] justifiedLines = SvgLines(justified, svgOptions);

        Assert.Equal(2, leftLines.Length);
        Assert.Equal(leftLines.Length, justifiedLines.Length);
        Assert.Equal(leftLines.Select(Words), justifiedLines.Select(Words));
        Assert.Equal("AA BB CC DD EE FF GG", string.Join(" ", justifiedLines.Select(Words)));
        Assert.InRange(Number(justifiedLines[0][0], "x") - Number(leftLines[0][0], "x"), -.002D, .002D);
        Assert.True(Number(justifiedLines[0].Last(), "x") > Number(leftLines[0].Last(), "x") + 3D,
            "The wrapped line must increase inter-word spacing.");
        Assert.InRange(Number(justifiedLines[1][0], "x") - Number(leftLines[1][0], "x"), -.002D, .002D);
        Assert.InRange(Number(justifiedLines[1].Last(), "x") - Number(leftLines[1].Last(), "x"), -.002D, .002D);

        var pngOptions = new VisioPngSaveOptions { Supersampling = 1, ResolveConnectorLabelOverlaps = false };
        pngOptions.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, font);
        (int Left, int Right)[] leftInk = RasterLineBounds(left.ToPng(pngOptions));
        (int Left, int Right)[] justifiedInk = RasterLineBounds(justified.ToPng(pngOptions));
        Assert.Equal(2, leftInk.Length);
        Assert.Equal(leftInk.Length, justifiedInk.Length);
        Assert.Equal(leftInk[0].Left, justifiedInk[0].Left);
        Assert.True(justifiedInk[0].Right > leftInk[0].Right + 3,
            "PNG must paint the widened spacing rather than merely retain the alignment metadata.");
        Assert.Equal(leftInk[1], justifiedInk[1]);
    }

    private static XElement[][] SvgLines(VisioPage page, VisioSvgSaveOptions options) =>
        XDocument.Parse(page.ToSvg(options)).Descendants(Svg + "g")
            .Single(node => (string?)node.Attribute("data-visio-shape-id") == "1"
                || (string?)node.Attribute("data-visio-connector-id") == "1")
            .Descendants(Svg + "text")
            .Where(node => !string.IsNullOrWhiteSpace(node.Value))
            .GroupBy(node => Number(node, "y"))
            .OrderBy(group => group.Key)
            .Select(group => group.ToArray()).ToArray();

    private static string Words(IEnumerable<XElement> line) =>
        string.Join(" ", line.Select(node => node.Value.Trim()));

    private static double Number(XElement node, string name) =>
        double.Parse(node.Attribute(name)!.Value, CultureInfo.InvariantCulture);

    private static (int Left, int Right)[] RasterLineBounds(byte[] png) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(png, out OfficeRasterImage? image));
        var lines = new List<(int Left, int Right)>();
        int left = image!.Width, right = -1;
        bool inLine = false;
        for (int y = 0; y < image.Height; y++) {
            bool painted = false;
            for (int x = 0; x < image.Width; x++) {
                OfficeColor color = image.GetPixel(x, y);
                bool red = color.R > 100 && color.R > color.G * 2 && color.R > color.B * 2;
                bool green = color.G > 50 && color.G > color.R * 2 && color.G > color.B * 1.5;
                if (!red && !green) continue;
                painted = true;
                left = Math.Min(left, x); right = Math.Max(right, x);
            }
            if (painted) inLine = true;
            else if (inLine) {
                lines.Add((left, right));
                left = image.Width; right = -1; inLine = false;
            }
        }
        if (inLine) lines.Add((left, right));
        return lines.ToArray();
    }

    private static VisioDocument NativeParagraph(bool connector, int alignment) {
        string transform = connector
            ? "<XForm1D><BeginX>1</BeginX><BeginY>1.5</BeginY><EndX>3</EndX><EndY>1.5</EndY></XForm1D>"
                + "<TextXForm><TxtWidth>1.3</TxtWidth><TxtHeight>1</TxtHeight></TextXForm>"
            : "<XForm><PinX>2</PinX><PinY>1.5</PinY><Width>1.3</Width><Height>1</Height></XForm>";
        string source = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'>"
            + $"<FaceNames><FaceName ID='0' Name='{ManagedTextShapingTestAssets.FamilyName}'/></FaceNames>"
            + "<Pages><Page ID='0' Name='Justified'><PageSheet><PageProps><PageWidth>4</PageWidth>"
            + "<PageHeight>3</PageHeight></PageProps></PageSheet><Shapes><Shape ID='1'>" + transform
            + "<Fill><FillPattern>0</FillPattern></Fill><Line><LinePattern>0</LinePattern></Line>"
            + "<TextBlock><LeftMargin>0</LeftMargin><RightMargin>0</RightMargin><TopMargin>0</TopMargin>"
            + "<BottomMargin>0</BottomMargin><VerticalAlign>0</VerticalAlign></TextBlock>"
            + "<Char IX='0'><Font>0</Font><Size>0.16666666666666667</Size><Color>#cc1122</Color></Char>"
            + "<Char IX='1'><Font>0</Font><Size>0.16666666666666667</Size><Color>#117733</Color></Char>"
            + $"<Para IX='0'><HorzAlign>{alignment}</HorzAlign><SpLine>-1.2</SpLine></Para>"
            + "<Text><pp IX='0'/><cp IX='0'/>AA <cp IX='1'/>BB <cp IX='0'/>CC <cp IX='1'/>DD "
            + "<cp IX='0'/>EE <cp IX='1'/>FF <cp IX='0'/>GG</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
    }
}

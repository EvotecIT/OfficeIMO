using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class VisioRichTextBulletTests {
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public void CustomLabelPreservesSizeFontAndHangingPositionWithoutChangingNativeText(bool connector, int mode) {
        var document = Reopen(Synthetic(connector, "<Bullet>1</Bullet><BulletStr>o.</BulletStr><BulletFont>5</BulletFont><BulletFontSize>0.25</BulletFontSize><TextPosAfterBullet>0.375</TextPosAfterBullet>"), mode);
        var page = document.Pages[0];
        byte[] before = document.ToLegacyXmlResult().Value;
        var text = PaintedText(page);
        Assert.Equal(new[] { "o.", "FIRST", "o.", "SECOND" }, text.Select(t => t.Value));
        Assert.Equal("24", (string?)text[0].Attribute("font-size"));
        Assert.Equal("Courier New", (string?)text[0].Attribute("font-family"));
        Assert.Equal("Arial", (string?)text[1].Attribute("font-family"));
        Assert.InRange(Number(text[1], "x") - Number(text[0], "x"), 35.99, 36.01);
        Assert.Equal("FIRST\nSECOND", connector ? page.Connectors[0].Label : page.Shapes[0].Text);
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }

    [Theory]
    [InlineData(1, "•")]
    [InlineData(2, "◆")]
    [InlineData(3, "■")]
    [InlineData(4, "□")]
    [InlineData(5, "❖")]
    [InlineData(6, "➢")]
    [InlineData(7, "✔")]
    public void StandardBulletKindsUseTheirUnicodeMarkersAndFirstCharacterFont(int kind, string marker) {
        foreach (bool connector in new[] { false, true }) {
            var page = Synthetic(connector, $"<Bullet>{kind}</Bullet><BulletFont>5</BulletFont><BulletFontSize>-1.5</BulletFontSize>", body: "FIRST").Pages[0];
            var text = PaintedText(page);
            Assert.Equal(marker, text[0].Value);
            Assert.Equal("24", (string?)text[0].Attribute("font-size"));
            Assert.Equal("Arial", (string?)text[0].Attribute("font-family"));
            Assert.InRange(Number(text[1], "x") - Number(text[0], "x"), 23.99, 24.01);

            var image = VisualBaselineTestSupport.DecodePng(page.ToPng(new VisioPngSaveOptions { Supersampling = 1 }), "Bullet PNG could not be decoded.");
            int bodyX = (int)Math.Ceiling(Number(text[1], "x"));
            int markerPixels = 0, bodyPixels = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.R <= 100 || pixel.G >= 90 || pixel.B >= 100) continue;
                    if (x < bodyX) markerPixels++;
                    else bodyPixels++;
                }
            }
            Assert.True(markerPixels > 3, $"Native bullet {kind} painted no marker pixels (connector: {connector}).");
            Assert.True(bodyPixels > 3, "The bullet body text was not painted.");
        }
    }

    [Theory]
    [InlineData("<BulletFontSize>0</BulletFontSize>", "16")]
    [InlineData("<BulletFontSize>-0.5</BulletFontSize>", "8")]
    public void CustomFontZeroAndRelativeSizesFollowTheFirstCharacter(string cells, string size) {
        var text = PaintedText(Synthetic(false, "<Bullet>1</Bullet><BulletStr>*</BulletStr><BulletFont>0</BulletFont>" + cells).Pages[0]).First();
        Assert.Equal("Arial", (string?)text.Attribute("font-family"));
        Assert.Equal(size, (string?)text.Attribute("font-size"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WrappedContinuationDoesNotRepeatTheBulletAndUsesNativeLeftIndent(bool connector) {
        var text = PaintedText(Synthetic(connector, "<Bullet>1</Bullet>", body: "FIRST ONE TWO THREE FOUR FIVE SIX SEVEN EIGHT NINE TEN ELEVEN TWELVE THIRTEEN").Pages[0]);
        Assert.Single(text, t => t.Value == "•");
        Assert.True(text.Length > 2);
        Assert.All(text.Skip(1), t => Assert.InRange(Number(t, "x") - Number(text[0], "x"), 23.99, 24.01));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EditingBulletRowsChangesBothRenderersAndLeavesCopiedRowsIndependent(bool connector) {
        var document = Synthetic(connector, "<Bullet>1</Bullet><BulletStr>*</BulletStr>");
        var page = document.Pages[0];
        var copy = document.DuplicatePage(page, "Copy");
        byte[] old = page.ToPng(new VisioPngSaveOptions { Supersampling = 1 });
        var section = (connector ? page.Connectors[0].GetShapeSheetSections() : page.Shapes[0].GetShapeSheetSections()).Single(s => s.Name is "Para" or "Paragraph");
        section.Rows.Single(r => r.Index == 9).SetCell("BulletStr", "++");
        if (connector) page.Connectors[0].SetShapeSheetSection(section);
        else page.Shapes[0].SetShapeSheetSection(section);
        Assert.Equal("++", Label(page).Descendants(Svg + "text").First().Value);
        Assert.Contains("*", copy.ToSvg());
        Assert.False(old.SequenceEqual(page.ToPng(new VisioPngSaveOptions { Supersampling = 1 })));
        Assert.Equal(old, copy.ToPng(new VisioPngSaveOptions { Supersampling = 1 }));
    }

    [Fact]
    public void LabelAndBodyShrinkTogetherInsideASmallNativeFrame() {
        var page = Synthetic(false, "<Bullet>1</Bullet><BulletStr>*</BulletStr><BulletFontSize>0.5</BulletFontSize>", height: .3, body: "FIRST").Pages[0];
        var text = PaintedText(page);
        Assert.Equal(new[] { "*", "FIRST" }, text.Select(t => t.Value));
        Assert.InRange(Number(text[0], "font-size"), 4, 24);
        Assert.InRange(Number(text[0], "font-size") / Number(text[1], "font-size"), 2.99, 3.01);
        Assert.InRange(Number(text[1], "x") - Number(text[0], "x"), 23.99, 24.01);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PlainLabelReplacementRemovesNativeBulletsAcrossReopening(bool connector) {
        var document = Synthetic(connector, "<Bullet>1</Bullet>");
        if (connector) document.Pages[0].Connectors[0].Label = "Replacement";
        else document.Pages[0].Shapes[0].Text = "Replacement";
        for (int mode = 0; mode < 3; mode++)
            Assert.Equal("Replacement", string.Concat(Label(Reopen(document, mode).Pages[0]).Descendants(Svg + "text").Select(t => t.Value)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BulletFontsReachShapingAndSubstitutionDiagnostics(bool connector) {
        var page = Synthetic(connector, "<Bullet>1</Bullet><BulletStr>++</BulletStr><BulletFont>Missing Bullet Font</BulletFont>").Pages[0];
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new VisioImageExportOptions { TextShapingProvider = provider, Supersampling = 1 };
        options.Fonts.Add("Missing Bullet Font", ManagedTextShapingTestAssets.CreateFont('+'));
        _ = page.ExportImage(OfficeImageExportFormat.Png, options);
        Assert.Contains(provider.Requests, r => r.Text.Contains("++"));
        Assert.Contains(page.ExportImage(OfficeImageExportFormat.Svg).Diagnostics, d => d.Message.Contains("Missing Bullet Font"));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void MasterLabelsUsePreservedBulletRowsWithoutMaterializingLocalText(int mode) {
        XNamespace native = "http://schemas.microsoft.com/visio/2003/core";
        var xml = XDocument.Parse(Encoding.UTF8.GetString(Synthetic(false, "<Bullet>1</Bullet><BulletStr>*</BulletStr>").ToLegacyXmlResult().Value));
        var shape = xml.Descendants(native + "Pages").Descendants(native + "Shape").Single();
        var master = new XElement(shape); master.SetAttributeValue("ID", "100");
        shape.ReplaceWith(new XElement(native + "Shape", new XAttribute("ID", "1"), new XAttribute("Master", "7"),
            new XElement(shape.Element(native + "XForm")!)));
        xml.Root!.AddFirst(new XElement(native + "Masters", new XElement(native + "Master", new XAttribute("ID", "7"), new XAttribute("NameU", "Bulleted"), new XElement(native + "Shapes", master))));
        var document = Reopen(VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)))).Value, mode);
        Assert.Equal(new[] { "*", "FIRST", "*", "SECOND" }, PaintedText(document.Pages[0]).Select(t => t.Value));
        var exported = XDocument.Parse(Encoding.UTF8.GetString(document.ToLegacyXmlResult().Value));
        Assert.Null(exported.Descendants(native + "Pages").Descendants(native + "Shape").Single().Element(native + "Text"));
    }

    [Fact]
    public void UnsupportedLabelsRetainCharacterProjectionAndDoNotThrowOnOversizedText() {
        foreach (string cells in new[] { "<Bullet>99</Bullet>", "<Bullet>1</Bullet><BulletStr>bad\nlabel</BulletStr>",
            "<Bullet>1</Bullet><BulletFontSize>NaN</BulletFontSize>", "<Bullet>1</Bullet><TextPosAfterBullet>-1</TextPosAfterBullet>",
            "<Bullet>1</Bullet><BulletStr>" + new string('*', 4097) + "</BulletStr>" }) {
            var page = Synthetic(false, cells, body: "FIRST", mixed: true).Pages[0];
            var text = PaintedText(page);
            Assert.Equal(new[] { "FIRST", "GREEN" }, text.Select(t => t.Value));
            Assert.Equal("#CC1122", (string?)text[0].Attribute("fill"));
            Assert.Equal("#117733", (string?)text[1].Attribute("fill"));
        }
        _ = Synthetic(false, "<Bullet>1</Bullet>", body: new string('a', 100000)).Pages[0].ToSvg();
    }

    internal static VisioDocument Synthetic(bool connector, string cells, double height = 2, string body = "FIRST\nSECOND", bool mixed = false) {
        string transform = connector
            ? $"<XForm1D><BeginX>1</BeginX><BeginY>2</BeginY><EndX>5</EndX><EndY>2</EndY></XForm1D><TextXForm><TxtWidth>4</TxtWidth><TxtHeight>{height.ToString(CultureInfo.InvariantCulture)}</TxtHeight></TextXForm>"
            : $"<XForm><PinX>3</PinX><PinY>2</PinY><Width>4</Width><Height>{height.ToString(CultureInfo.InvariantCulture)}</Height></XForm>";
        string source = $"<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><FaceNames><FaceName ID='0' Name='Unselected Zero Font'/><FaceName ID='3' Name='Arial'/><FaceName ID='5' Name='Courier New'/></FaceNames><Pages><Page ID='0'><PageSheet><PageProps><PageWidth>6</PageWidth><PageHeight>4</PageHeight></PageProps></PageSheet><Shapes><Shape ID='1'>{transform}<Fill><FillPattern>0</FillPattern></Fill><Line><LinePattern>0</LinePattern></Line><TextBlock><VerticalAlign>0</VerticalAlign></TextBlock><Char IX='3'><Font>3</Font><Color>#cc1122</Color><Size>0.16666666666666667</Size></Char><Char IX='4'><Font>3</Font><Color>#117733</Color><Size>0.16666666666666667</Size></Char><Para IX='9'><IndFirst>-0.25</IndFirst><IndLeft>0.25</IndLeft><SpLine>-1.2</SpLine><HorzAlign>0</HorzAlign>{cells}</Para><Text><cp IX='3'/><pp IX='9'/>{body}{(mixed ? "<cp IX='4'/>GREEN" : "")}</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
    }

    private static VisioDocument Reopen(VisioDocument document, int mode) => mode switch {
        1 => VisioDocument.Load(new MemoryStream(document.ToBytes())),
        2 => VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value,
        _ => document
    };
    private static XElement Label(VisioPage page) => XDocument.Parse(page.ToSvg()).Descendants(Svg + "g")
        .Single(t => (string?)t.Attribute("data-visio-shape-id") == "1" || (string?)t.Attribute("data-visio-connector-id") == "1");
    private static double Number(XElement node, string name) => double.Parse(node.Attribute(name)!.Value, CultureInfo.InvariantCulture);
    private static XElement[] PaintedText(VisioPage page) => Label(page).Descendants(Svg + "text").Where(t => !string.IsNullOrWhiteSpace(t.Value)).ToArray();
}

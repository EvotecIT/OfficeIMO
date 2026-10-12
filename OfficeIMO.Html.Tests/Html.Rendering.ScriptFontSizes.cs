using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using UglyToad.PdfPig.Core;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("sup", "font-size:75%", 12D, OfficeTextBaseline.Superscript)]
    [InlineData("sub", "font-size:75%", 12D, OfficeTextBaseline.Subscript)]
    [InlineData("span", "vertical-align:super", 16D, OfficeTextBaseline.Superscript)]
    [InlineData("span", "vertical-align:sub", 16D, OfficeTextBaseline.Subscript)]
    [InlineData("sup", "", 16D / 1.2D, OfficeTextBaseline.Superscript)]
    [InlineData("sub", "", 16D / 1.2D, OfficeTextBaseline.Subscript)]
    [InlineData("sup", "font-size:smaller", 16D / 1.2D, OfficeTextBaseline.Superscript)]
    [InlineData("sup", "font-size:inherit", 16D, OfficeTextBaseline.Superscript)]
    [InlineData("sup", "font-size:initial", 16D, OfficeTextBaseline.Superscript)]
    [InlineData("sup", "font-size:revert", 16D / 1.2D, OfficeTextBaseline.Superscript)]
    [InlineData("sup", "font-size:75%;vertical-align:baseline", 12D, OfficeTextBaseline.Normal)]
    [InlineData("sub", "font-size:75%;vertical-align:initial", 12D, OfficeTextBaseline.Normal)]
    [InlineData("sup", "font-size:75%;vertical-align:unset", 12D, OfficeTextBaseline.Normal)]
    [InlineData("sup", "font-size:75%;vertical-align:sub", 12D, OfficeTextBaseline.Subscript)]
    public void HtmlScriptText_PaintsTheResolvedCssFontSizeWithoutAdditionalScriptScaling(
        string tag, string style, double expectedSize, OfficeTextBaseline expectedBaseline) {
        string html = "<p style='margin:0;font:16px/32px Helvetica'>BASE <" + tag
            + " style='" + style + "'>X</" + tag + "></p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var options = new HtmlRenderOptions { Margins = HtmlRenderMargins.All(0D) };
        HtmlRenderText text = Assert.Single(
            EnumerateRenderVisuals(HtmlRenderTestDriver.Render(document, options).Pages[0].Scene)
                .OfType<HtmlRenderText>(), item => item.Text == "X");
        Assert.Equal(expectedSize, text.Font.Size, 6);
        Assert.Equal(1D, text.BaselineScale);
        Assert.Equal(expectedBaseline, text.Baseline);

        XElement painted = Assert.Single(XDocument.Parse(document.ToSvg(options)).Descendants(),
            node => node.Name.LocalName == "text" && node.Value == "X");
        Assert.Equal(expectedSize, ParseScriptCoordinate(painted, "font-size"), 3);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options)));
        var visible = pdf.GetPage(1).Letters.Where(letter => letter.RenderingMode == TextRenderingMode.Fill).ToArray();
        var pdfX = Assert.Single(visible, letter => letter.Value == "X");
        Assert.Equal(expectedSize * 0.75D, pdfX.FontSize, 3);
    }

    [Fact]
    public void HtmlScriptText_DescendantFontOverridesDoNotInheritPaintScaling() {
        const string html = "<p style='margin:0;font:24px/40px Helvetica'>"
            + "<sup style='font-size:75%'><span>Inherited</span>"
            + "<span style='font-size:20px'>Override</span>"
            + "<sup style='font-size:50%'>Nested</sup>"
            + "<span style='vertical-align:baseline'>ParentBaseline</span></sup>"
            + "<sup><sup>DefaultNested</sup></sup></p>";
        HtmlRenderText[] text = EnumerateRenderVisuals(HtmlRenderTestDriver.Render(html).Pages[0].Scene)
            .OfType<HtmlRenderText>().ToArray();
        HtmlRenderText inherited = Assert.Single(text, item => item.Text == "Inherited");
        HtmlRenderText overridden = Assert.Single(text, item => item.Text == "Override");
        HtmlRenderText nested = Assert.Single(text, item => item.Text == "Nested");
        HtmlRenderText parentBaseline = Assert.Single(text, item => item.Text == "ParentBaseline");
        HtmlRenderText defaultNested = Assert.Single(text, item => item.Text == "DefaultNested");
        Assert.Equal(18D, inherited.Font.Size);
        Assert.Equal(20D, overridden.Font.Size);
        Assert.Equal(9D, nested.Font.Size);
        Assert.Equal(24D / 1.2D / 1.2D, defaultNested.Font.Size, 6);
        Assert.All(new[] { inherited, overridden, nested, parentBaseline, defaultNested },
            item => Assert.Equal(1D, item.BaselineScale));
        Assert.Equal(inherited.BaselineOffset, overridden.BaselineOffset);
        Assert.Equal(inherited.BaselineOffset, parentBaseline.BaselineOffset);
        Assert.True(nested.BaselineOffset < inherited.BaselineOffset);
    }

    [Theory]
    [InlineData("sup")]
    [InlineData("sub")]
    public void HtmlScriptText_FontSizeControlsAdvancesWrappingAndFollowingBlock(string tag) {
        string actualHtml = "<div style='width:100px;font:16px/40px Helvetica'><" + tag
            + " style='font-size:12px'>MMMMMMMM</" + tag + "> <span>TAIL</span></div><div>FOLLOWING</div>";
        const string controlHtml = "<div style='width:100px;font:16px/40px Helvetica'>"
            + "<span style='font-size:12px'>MMMMMMMM</span> <span>TAIL</span></div><div>FOLLOWING</div>";
        HtmlRenderText[] actual = Text(actualHtml);
        HtmlRenderText[] control = Text(controlHtml);
        HtmlRenderText script = Assert.Single(actual, item => item.Text == "MMMMMMMM");
        HtmlRenderText plain = Assert.Single(control, item => item.Text == "MMMMMMMM");
        Assert.Equal(plain.TextAdvanceWidth!.Value, script.TextAdvanceWidth!.Value, 6);
        Assert.True(Assert.Single(actual, item => item.Text == "TAIL").LayoutY > script.LayoutY);
        Assert.Equal(Assert.Single(control, item => item.Text == "TAIL").LayoutY,
            Assert.Single(actual, item => item.Text == "TAIL").LayoutY, 6);
        Assert.Equal(Assert.Single(control, item => item.Text == "FOLLOWING").LayoutY,
            Assert.Single(actual, item => item.Text == "FOLLOWING").LayoutY, 6);

        static HtmlRenderText[] Text(string html) =>
            EnumerateRenderVisuals(HtmlRenderTestDriver.Render(html).Pages[0].Scene).OfType<HtmlRenderText>().ToArray();
    }

    [Theory]
    [InlineData("50%")]
    [InlineData(".5em")]
    public void HtmlScriptText_NumericShiftUsesTheAlignedElementsOwnFontAndLineHeight(string alignment) {
        string html = "<p style='margin:0;font:20px/40px Helvetica'>BASE "
            + "<span style='font-size:10px;line-height:10px;vertical-align:" + alignment + "'>X</span></p>";
        HtmlRenderText text = Assert.Single(
            EnumerateRenderVisuals(HtmlRenderTestDriver.Render(html).Pages[0].Scene).OfType<HtmlRenderText>(),
            item => item.Text == "X");
        Assert.Equal(10D, text.Font.Size);
        Assert.Equal(-5D, text.BaselineOffset, 6);
    }

    [Theory]
    [InlineData("", 11.2D, true)]
    [InlineData("font-size:70%", 11.2D, true)]
    [InlineData("font-size:revert", 11.2D, true)]
    [InlineData("font-size:16px", 16D, true)]
    [InlineData("font-size:16px;vertical-align:baseline", 16D, false)]
    [InlineData("font-size:initial;vertical-align:initial", 16D, false)]
    [InlineData("font-size:inherit;vertical-align:baseline", 16D, false)]
    [InlineData("font-size:unset;vertical-align:baseline", 16D, false)]
    [InlineData("font-size:16px;vertical-align:unset", 16D, false)]
    public void HtmlFootnoteCall_PaintsItsDefaultOrAuthoredFontOnceAndRaisesItsBaseline(
        string callStyle, double expectedSize, bool raised) {
        string html = "<style>body{margin:0;font:16px/24px Helvetica}.note{float:footnote}"
            + ".note::footnote-marker{content:none}.note::footnote-call{" + callStyle + "}</style>"
            + "<p style='margin:0'>BASE<span class='note'>Note text</span> AFTER</p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, Margins = HtmlRenderMargins.All(0D) };
        HtmlRenderText call = Assert.Single(
            EnumerateRenderVisuals(HtmlRenderTestDriver.Render(document, options).Pages[0].Scene).OfType<HtmlRenderText>(),
            item => item.SemanticRole == "footnote-call" || item.Source?.EndsWith(":footnote-call", StringComparison.Ordinal) == true);
        Assert.Equal(expectedSize, call.Font.Size, 6);
        Assert.Equal(1D, call.BaselineScale);
        Assert.Equal(raised ? OfficeTextBaseline.Superscript : OfficeTextBaseline.Normal, call.Baseline);
        Assert.Equal(raised, call.BaselineOffset < 0D);

        XElement[] svgText = XDocument.Parse(document.ToSvg(options)).Descendants()
            .Where(node => node.Name.LocalName == "text").ToArray();
        XElement svgCall = Assert.Single(svgText, node => node.Value == "1");
        XElement svgBase = Assert.Single(svgText, node => node.Value == "BASE");
        Assert.Equal(expectedSize, ParseScriptCoordinate(svgCall, "font-size"), 3);
        if (raised) Assert.True(ParseScriptCoordinate(svgCall, "y") < ParseScriptCoordinate(svgBase, "y"));
        else Assert.Equal(ParseScriptCoordinate(svgBase, "y"), ParseScriptCoordinate(svgCall, "y"), 3);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options)));
        var visible = pdf.GetPage(1).Letters.Where(letter => letter.RenderingMode == TextRenderingMode.Fill).ToArray();
        var pdfCall = Assert.Single(visible, letter => letter.Value == "1");
        Assert.Equal(expectedSize * 0.75D, pdfCall.FontSize, 3);
        var pdfBase = Assert.Single(visible, letter => letter.Value == "B");
        if (raised) Assert.True(pdfCall.StartBaseLine.Y > pdfBase.StartBaseLine.Y);
        else Assert.Equal(pdfBase.StartBaseLine.Y, pdfCall.StartBaseLine.Y, 3);
    }

    private static double ParseScriptCoordinate(XElement node, string attribute) =>
        double.Parse(node.Attribute(attribute)!.Value, CultureInfo.InvariantCulture);
}

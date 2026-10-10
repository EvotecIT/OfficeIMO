using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlInlineDecorationGeometryTests {
    [Theory]
    [InlineData("top")]
    [InlineData("baseline")]
    public void ParentDecorationUsesItsFontBoxWithoutChangingTheAtomicChild(string alignment) {
        HtmlRenderOptions options = Options();
        HtmlRenderDocument rendered = Render("<div style='height:120px;padding:10px;border:2px solid red'>"
            + "<span id='owner' style='height:20px;padding:7px;border:1px solid black;background:yellow;box-sizing:border-box'>"
            + Atom("fill", 120D, alignment) + "</span></div>" + Following, options);
        HtmlRenderShape owner = Shape(rendered, "span#owner");
        HtmlRenderShape fill = Shape(rendered, "span#fill");
        var face = Face(options);

        Assert.Equal(face.Height + 16D, owner.Height, 6);
        Assert.Equal(56D, owner.Width, 6);
        Assert.Equal(12D, owner.X, 6);
        Assert.Equal(20D, fill.X, 6);
        Assert.Equal(12D, fill.Y, 6);
        Assert.Equal(120D, fill.Height, 6);
        Assert.Equal(144D, Shape(rendered, "div#following").Y, 6);
        if (alignment == "top") Assert.InRange(owner.Y, 4D, 5D);
        else Assert.Equal(fill.Y + fill.Height - face.Ascent - 8D, owner.Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("")]
    [InlineData("Text")]
    public void TextAndEmptyInlineDecorationIgnoreATallSibling(string text) {
        HtmlRenderOptions options = Options();
        HtmlRenderDocument rendered = Render("<div><span id='owner' style='padding:7px;border:1px solid black;background:yellow'>"
            + text + "</span>" + Atom("fill", 120D, "top") + "</div>" + Following, options);

        Assert.Equal(Face(options).Height + 16D, Shape(rendered, "span#owner").Height, 6);
        Assert.Equal(120D, Shape(rendered, "span#fill").Height, 6);
        Assert.Equal(120D, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void NestedDecorationKeepsEachOwnersInsetsAndTheSharedFontHeight() {
        HtmlRenderOptions options = Options();
        HtmlRenderDocument rendered = Render("<div><span id='outer' style='padding:7px;border:1px solid black;background:yellow'>"
            + "<span id='inner' style='padding:3px;border:2px solid red;background:blue'>"
            + Atom("fill", 120D, "top") + "</span></span></div>", options);
        HtmlRenderShape outer = Shape(rendered, "span#outer");
        HtmlRenderShape inner = Shape(rendered, "span#inner");

        Assert.Equal(Face(options).Height + 16D, outer.Height, 6);
        Assert.Equal(Face(options).Height + 10D, inner.Height, 6);
        Assert.Equal(66D, outer.Width, 6);
        Assert.Equal(50D, inner.Width, 6);
        Assert.Equal(8D, inner.X - outer.X, 6);
        Assert.Equal(3D, inner.Y - outer.Y, 6);
        Assert.Equal(13D, Shape(rendered, "span#fill").X, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("slice", 0D, 48D)]
    [InlineData("clone", 8D, 56D)]
    public void WrappedDecorationUsesAFontBoxOnEachLineAndRetainsItsEdges(string decorationBreak, double secondX, double fragmentWidth) {
        HtmlRenderOptions options = Options();
        HtmlRenderDocument rendered = Render("<div style='width:70px'><span id='owner' style='padding:7px;border:1px solid black;"
            + "background:yellow;box-decoration-break:" + decorationBreak + "'>"
            + Atom("first", 120D, "top") + Atom("second", 60D, "top") + "</span></div>" + Following, options);
        HtmlRenderShape[] fragments = Shapes(rendered, "span#owner").OrderBy(shape => shape.Y).ToArray();

        Assert.Equal(2, fragments.Length);
        Assert.All(fragments, fragment => {
            Assert.Equal(Face(options).Height + 16D, fragment.Height, 6);
            Assert.Equal(fragmentWidth, fragment.Width, 6);
        });
        Assert.Equal(120D, fragments[1].Y - fragments[0].Y, 6);
        Assert.Equal(8D, Shape(rendered, "span#first").X, 6);
        Assert.Equal(secondX, Shape(rendered, "span#second").X, 6);
        Assert.Equal(180D, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void ADescendantFloatDoesNotSupplyTheInlineDecorationExtent() {
        HtmlRenderOptions options = Options();
        const string opening = "<div><span id='owner' style='padding:7px;border:1px solid black;background:yellow'>";
        const string closing = "Text</span></div>";
        HtmlRenderDocument plain = Render(opening + closing, options);
        HtmlRenderDocument floated = Render(opening
            + "<span id='float' style='float:left;width:40px;height:120px;background:lime'></span>" + closing, options);

        Assert.Equal(Shape(plain, "span#owner").Width, Shape(floated, "span#owner").Width, 6);
        Assert.Equal(Face(options).Height + 16D, Shape(floated, "span#owner").Height, 6);
        Assert.Equal(120D, Shape(floated, "span#float").Height, 6);
        floated.RequireNoLoss();
    }

    [Fact]
    public void FontDecorationDoesNotMoveTheExistingPositionedChild() {
        HtmlRenderOptions options = Options();
        string Source(string paint) => "<div><span id='owner' style='position:relative;padding:7px;"
            + "border:1px solid " + paint + "'>" + Atom("fill", 120D, "top")
            + "<span id='positioned' style='position:absolute;left:0;top:100%;width:10px;height:10px;background:blue'></span>"
            + "</span></div>" + Following;
        HtmlRenderDocument control = Render(Source("transparent"), options);
        HtmlRenderDocument rendered = Render(Source("black;background:yellow"), options);
        HtmlRenderShape positioned = Shape(rendered, "span#positioned");
        HtmlRenderShape controlPositioned = Shape(control, "span#positioned");

        Assert.Equal(Face(options).Height + 16D, Shape(rendered, "span#owner").Height, 6);
        Assert.Equal(controlPositioned.X, positioned.X, 6);
        Assert.Equal(controlPositioned.Y, positioned.Y, 6);
        Assert.Equal(Shape(control, "div#following").Y, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(0D, "text")]
    [InlineData(8D, "text")]
    [InlineData(20D, "text")]
    [InlineData(0D, "text-empty-sibling")]
    [InlineData(8D, "text-empty-sibling")]
    [InlineData(20D, "text-empty-sibling")]
    [InlineData(0D, "empty-standalone")]
    [InlineData(8D, "empty-standalone")]
    [InlineData(20D, "empty-standalone")]
    public void DecorationSharesTheTextOrEmptyStrutBaselineWithShortAndNormalLines(double lineHeight, string content) {
        HtmlRenderOptions options = Options();
        string Source(bool paint) {
            string edges = "padding:7px;border:1px solid " + (paint ? "black;background:yellow" : "transparent");
            string inline = content == "empty-standalone" ? string.Empty
                : "<span id='text-owner' style='" + edges + "'>Text</span>";
            if (content != "text") inline += "<span id='empty-owner' style='" + edges + "'></span>";
            return "<div style='margin-top:30px;line-height:"
                + lineHeight.ToString(System.Globalization.CultureInfo.InvariantCulture) + "px'>"
                + inline + "</div>" + Following;
        }
        HtmlRenderDocument control = Render(Source(false), options);
        HtmlRenderDocument rendered = Render(Source(true), options);
        var face = Face(options);
        double usedLineHeight = Shape(control, "div#following").Y - 30D;
        double selectedBaseline = lineHeight < 16D ? (usedLineHeight - face.Height) / 2D + face.Ascent : 16D;
        double expectedTop = 30D + selectedBaseline - face.Ascent - 8D;

        if (content != "empty-standalone") {
            HtmlRenderText text = Assert.Single(rendered.Pages.SelectMany(page => Enumerate(page.Scene))
                .OfType<HtmlRenderText>().Where(text => text.Text == "Text"));
            HtmlRenderText controlText = Assert.Single(control.Pages.SelectMany(page => Enumerate(page.Scene))
                .OfType<HtmlRenderText>().Where(text => text.Text == "Text"));
            Assert.Equal(text.Y + text.Font.Size - face.Ascent - 8D,
                Shape(rendered, "span#text-owner").Y, 6);
            Assert.Equal(expectedTop, Shape(rendered, "span#text-owner").Y, 6);
            Assert.Equal(face.Height + 16D, Shape(rendered, "span#text-owner").Height, 6);
            Assert.Equal(controlText.X, text.X, 6);
            Assert.Equal(controlText.Y, text.Y, 6);
        }
        if (content != "text") {
            Assert.Equal(expectedTop, Shape(rendered, "span#empty-owner").Y, 6);
            Assert.Equal(face.Height + 16D, Shape(rendered, "span#empty-owner").Height, 6);
        }
        Assert.Equal(Shape(control, "div#following").Y, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(8D, "X")]
    [InlineData(20D, "X")]
    [InlineData(8D, ".")]
    [InlineData(20D, ".")]
    [InlineData(8D, " ")]
    [InlineData(20D, " ")]
    public void GeneratedLeadersRetainTheirSpecializedDecorationAndPaintPlacement(double lineHeight, string pattern) {
        HtmlRenderOptions options = Options();
        string Source(bool paint) => "<style>#owner::before{content:leader(\"" + pattern + "\")}</style>"
            + "<div style='margin-top:30px;width:120px;line-height:"
            + lineHeight.ToString(System.Globalization.CultureInfo.InvariantCulture) + "px'>"
            + "<span id='owner' style='padding:7px;border:1px solid "
            + (paint ? "black;background:yellow" : "transparent") + "'>"
            + (pattern == " " ? "Text" : string.Empty) + "</span></div>" + Following;
        HtmlRenderDocument control = Render(Source(false), options);
        HtmlRenderDocument rendered = Render(Source(true), options);
        // This is the existing specialized leader route, outside the ordinary
        // text/empty-strut font-box policy. Its glyph/stroke and flow stay put.
        HtmlRenderShape[] owners = Shapes(rendered, "span#owner").ToArray();
        if (pattern == " " && lineHeight < 16D) {
            HtmlRenderText text = Assert.Single(rendered.Pages.SelectMany(page => Enumerate(page.Scene))
                .OfType<HtmlRenderText>(), text => text.Text == "Text");
            Assert.Equal(2, owners.Length);
            Assert.Contains(owners, owner => Math.Abs(owner.Y - 22D) < 0.000001D
                && Math.Abs(owner.Height - lineHeight - 16D) < 0.000001D);
            Assert.Contains(owners, owner => Math.Abs(owner.Y - text.Y + 8D) < 0.000001D
                && Math.Abs(owner.Height - text.Height - 16D) < 0.000001D);
        } else {
            HtmlRenderShape owner = Assert.Single(owners);
            Assert.Equal(22D, owner.Y, 6);
            Assert.Equal(lineHeight + 16D, owner.Height, 6);
        }
        Assert.Equal(Shape(control, "div#following").Y, Shape(rendered, "div#following").Y, 6);
        HtmlRenderVisual[] LeaderPaint(HtmlRenderDocument document) => document.Pages
            .SelectMany(page => Enumerate(page.Scene))
            .Where(visual => visual.Source?.Contains(":content-leader", StringComparison.Ordinal) == true)
            .ToArray();
        HtmlRenderVisual[] controlLeader = LeaderPaint(control);
        HtmlRenderVisual[] leader = LeaderPaint(rendered);
        Assert.Equal(controlLeader.Length, leader.Length);
        for (int index = 0; index < leader.Length; index++) {
            Assert.Equal(controlLeader[index].GetType(), leader[index].GetType());
            Assert.Equal(controlLeader[index].X, leader[index].X, 6);
            Assert.Equal(controlLeader[index].Y, leader[index].Y, 6);
            Assert.Equal(controlLeader[index].Width, leader[index].Width, 6);
            Assert.Equal(controlLeader[index].Height, leader[index].Height, 6);
        }
        if (pattern == "X") {
            HtmlRenderText text = Assert.Single(leader.OfType<HtmlRenderText>());
            Assert.Equal(30D, text.Y, 6);
            Assert.Equal(16D, text.Font.Size, 6);
        } else if (pattern == ".") {
            Assert.Single(leader.OfType<HtmlRenderShape>());
        } else {
            Assert.Empty(leader);
            HtmlRenderText text = Assert.Single(rendered.Pages.SelectMany(page => Enumerate(page.Scene))
                .OfType<HtmlRenderText>(), text => text.Text == "Text");
            HtmlRenderText controlText = Assert.Single(control.Pages.SelectMany(page => Enumerate(page.Scene))
                .OfType<HtmlRenderText>(), text => text.Text == "Text");
            Assert.Equal(controlText.X, text.X, 6);
            Assert.Equal(controlText.Y, text.Y, 6);
        }
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(8D)]
    [InlineData(20D)]
    public void SpaceOnlyLeaderPaintsItsOwnerWithoutGlyphsAndRespectsVisibility(double lineHeight) {
        HtmlRenderOptions options = Options();
        string Source(bool visible) => "<style>#owner::before{content:leader(\" \")}</style>"
            + "<div style='margin-top:30px;width:120px;line-height:"
            + lineHeight.ToString(System.Globalization.CultureInfo.InvariantCulture) + "px'>"
            + "<span id='owner' style='padding:7px;border:1px solid black;background:yellow;visibility:"
            + (visible ? "visible" : "hidden") + "'></span></div>" + Following;
        HtmlRenderDocument hidden = Render(Source(false), options);
        HtmlRenderDocument rendered = Render(Source(true), options);
        HtmlRenderShape owner = Shape(rendered, "span#owner");

        Assert.Equal(120D, owner.Width, 6);
        Assert.Equal(22D, owner.Y, 6);
        Assert.Equal(lineHeight + 16D, owner.Height, 6);
        Assert.Empty(Shapes(hidden, "span#owner"));
        Assert.DoesNotContain(rendered.Pages.SelectMany(page => Enumerate(page.Scene)), visual =>
            visual.Source?.Contains(":content-leader", StringComparison.Ordinal) == true);
        Assert.Equal(Shape(hidden, "div#following").Y, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(8D)]
    [InlineData(20D)]
    public void HiddenTextRegistersItsVisibleAncestorsDecorationWithoutExposingGlyphsOrLinks(double lineHeight) {
        HtmlRenderOptions options = Options();
        string hiddenText = VisibilitySource(lineHeight, true, false);
        HtmlRenderDocument rendered = Render(hiddenText, options);
        HtmlRenderDocument visibleText = Render(VisibilitySource(lineHeight, true, true), options);
        HtmlRenderDocument hiddenOwner = Render(VisibilitySource(lineHeight, false, false), options);
        HtmlRenderShape owner = Shape(rendered, "span#owner");
        HtmlRenderShape control = Shape(visibleText, "span#owner");

        Assert.Equal(control.X, owner.X, 6);
        Assert.Equal(control.Y, owner.Y, 6);
        Assert.Equal(control.Width, owner.Width, 6);
        Assert.Equal(control.Height, owner.Height, 6);
        Assert.Empty(Shapes(hiddenOwner, "span#owner"));
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => Enumerate(page.Scene)).ToArray();
        Assert.Empty(visuals.OfType<HtmlRenderText>());
        Assert.DoesNotContain(visuals, visual => visual.LinkUri != null);
        Assert.Equal(Shape(visibleText, "div#following").Y, Shape(rendered, "div#following").Y, 6);
        Assert.Equal(Shape(hiddenOwner, "div#following").Y, Shape(rendered, "div#following").Y, 6);
        var pdf = PdfReadDocument.Open(HtmlConversionDocument.Parse(WithPinnedStyle(hiddenText))
            .RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
                HtmlRenderEncoder.Pdf, options: options)).ToBytes());
        Assert.Empty(pdf.ExtractText());
        Assert.Empty(pdf.Pages.SelectMany(page => page.GetLinkAnnotations()));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(8D)]
    [InlineData(20D)]
    public void AVisibleTextOverrideDoesNotRevealItsHiddenAncestorsDecoration(double lineHeight) {
        HtmlRenderOptions options = Options();
        HtmlRenderDocument rendered = Render(VisibilitySource(lineHeight, false, true), options);
        HtmlRenderDocument control = Render(VisibilitySource(lineHeight, true, true), options);
        HtmlRenderText Text(HtmlRenderDocument document) => Assert.Single(document.Pages
            .SelectMany(page => Enumerate(page.Scene)).OfType<HtmlRenderText>(), text => text.Text == "Text");
        HtmlRenderText text = Text(rendered);

        Assert.Empty(Shapes(rendered, "span#owner"));
        Assert.Equal(Text(control).X, text.X, 6);
        Assert.Equal(Text(control).Y, text.Y, 6);
        Assert.Equal("https://example.com/visible", text.LinkUri);
        Assert.Equal(Shape(control, "div#following").Y, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HiddenAtomicContentRegistersItsVisibleAncestorsDecorationWithoutPaintingTheChild(bool image) {
        HtmlRenderOptions options = Options();
        string Source(bool ownerVisible, bool childVisible) {
            string style = "width:40px;height:" + (image ? "20" : "120")
                + "px;vertical-align:top;visibility:" + (childVisible ? "visible" : "hidden");
            string child = image
                ? "<img id='child' alt='' style='" + style + "' src='data:image/png;base64,"
                    + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2)) + "'>"
                : "<span id='child' style='display:inline-block;background:lime;" + style + "'></span>";
            return "<div style='margin-top:30px;width:120px'><span id='owner' "
                + "style='padding:7px;border:1px solid black;background:yellow;visibility:"
                + (ownerVisible ? "visible" : "hidden") + "'>" + child + "</span></div>" + Following;
        }
        HtmlRenderDocument rendered = Render(Source(true, false), options);
        HtmlRenderDocument visibleChild = Render(Source(true, true), options);
        HtmlRenderDocument hiddenOwner = Render(Source(false, false), options);
        HtmlRenderShape owner = Shape(rendered, "span#owner");
        HtmlRenderShape control = Shape(visibleChild, "span#owner");

        Assert.Equal(control.X, owner.X, 6);
        Assert.Equal(control.Y, owner.Y, 6);
        Assert.Equal(control.Width, owner.Width, 6);
        Assert.Equal(control.Height, owner.Height, 6);
        Assert.Empty(Shapes(hiddenOwner, "span#owner"));
        Assert.Empty(Shapes(rendered, "span#child"));
        Assert.Empty(rendered.Pages.SelectMany(page => Enumerate(page.Scene)).OfType<HtmlRenderImage>());
        Assert.Equal(Shape(visibleChild, "div#following").Y, Shape(rendered, "div#following").Y, 6);
        Assert.Equal(Shape(hiddenOwner, "div#following").Y, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(8D)]
    [InlineData(20D)]
    public void AZeroShareLeaderPaintsItsOwnersEdgesWithoutCreatingLeaderVisuals(double lineHeight) {
        HtmlRenderOptions options = Options();
        string Source(bool decorated) => "<style>#owner::before{content:leader(\"X\")}</style>"
            + "<div style='margin-top:30px;width:16px;line-height:"
            + lineHeight.ToString(System.Globalization.CultureInfo.InvariantCulture) + "px'>"
            + "<span id='owner' style='padding:7px;border:1px solid "
            + (decorated ? "black;background:yellow" : "transparent;background:transparent")
            + "'></span></div>" + Following;
        HtmlRenderDocument rendered = Render(Source(true), options);
        HtmlRenderDocument control = Render(Source(false), options);
        HtmlRenderShape owner = Shape(rendered, "span#owner");

        Assert.InRange(owner.Width, 16D, 16.011D);
        Assert.Equal(22D, owner.Y, 6);
        Assert.Equal(lineHeight + 16D, owner.Height, 6);
        Assert.DoesNotContain(rendered.Pages.SelectMany(page => Enumerate(page.Scene)), visual =>
            visual.Source?.Contains(":content-leader", StringComparison.Ordinal) == true);
        Assert.Equal(Shape(control, "div#following").Y, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    private static string VisibilitySource(double lineHeight, bool ownerVisible, bool childVisible) =>
        "<div style='margin-top:30px;width:120px;line-height:"
        + lineHeight.ToString(System.Globalization.CultureInfo.InvariantCulture) + "px'>"
        + "<span id='owner' style='padding:7px;border:1px solid black;background:yellow;visibility:"
        + (ownerVisible ? "visible" : "hidden") + "'><a href='https://example.com/"
        + (childVisible ? "visible" : "hidden") + "' style='visibility:"
        + (childVisible ? "visible" : "hidden") + ";text-decoration:none;color:black'>Text</a></span></div>" + Following;

    private static string WithPinnedStyle(string content) =>
        "<!doctype html><style>html,body{margin:0;padding:0;font:16px/20px Pinned}</style>" + content;

    private const string Following = "<div id='following' style='height:10px;background:blue'></div>";
    private static string Atom(string id, double height, string alignment) => "<span id='" + id
        + "' style='display:inline-block;width:40px;height:" + height.ToString(System.Globalization.CultureInfo.InvariantCulture)
        + "px;vertical-align:" + alignment + ";background:lime'></span>";

    private static HtmlRenderOptions Options() {
        var options = new HtmlRenderOptions { ViewportWidth = 600D, ViewportHeight = 900D,
            PageSize = new OfficePageSize(600D / 96D, 900D / 96D), Margins = HtmlRenderMargins.All(0D),
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser, HonorCssPageRules = false };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
        return options;
    }

    private static (double Height, double Ascent) Face(HtmlRenderOptions options) {
        IOfficeFontProgram program = options.Fonts.ResolveForText(string.Empty, "Pinned",
            OfficeFontFaceDescriptor.Regular, 16D, out _)!;
        return (program.LineHeight(16D), Assert.IsAssignableFrom<IOfficeFontBaselineMetrics>(program).BaselineOffset(16D));
    }

    private static HtmlRenderDocument Render(string content, HtmlRenderOptions options) => HtmlRenderEngine.Execute(
        HtmlConversionDocument.Parse(WithPinnedStyle(content)),
        HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: options)).Document;

    private static IEnumerable<HtmlRenderShape> Shapes(HtmlRenderDocument document, string source) =>
        document.Pages.SelectMany(page => Enumerate(page.Scene)).OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == source && shape.Shape.FillColor.HasValue);
    private static HtmlRenderShape Shape(HtmlRenderDocument document, string source) => Assert.Single(Shapes(document, source));
    private static IEnumerable<HtmlRenderVisual> Enumerate(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderSemanticGroup group => group.Visuals,
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                _ => null
            };
            if (children != null) foreach (HtmlRenderVisual child in Enumerate(children)) yield return child;
        }
    }
}

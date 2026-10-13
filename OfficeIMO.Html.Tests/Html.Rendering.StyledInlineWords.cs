using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlStyledWord_NoSpaceRunsOverflowTogetherAndRetainTheirIntrinsicWidth(bool script) {
        var calibration = CalibrateStyledWord(script);
        HtmlRenderDocument rendered = RenderStyledWord(StyledWordMarkup(script), StyledWordWidth(calibration.Width));
        HtmlRenderText left = StyledWordText(rendered, script ? "steradian" : "stera");
        HtmlRenderText right = StyledWordText(rendered, script ? "(b)" : "dian");

        Assert.Equal(calibration.Right.Y - calibration.Left.Y, right.Y - left.Y, 6);
        Assert.Equal(left.X + StyledWordAdvance(left), right.X, 6);
        Assert.Equal(StyledWordBox(calibration.Document).Height, StyledWordBox(rendered).Height, 6);
        Assert.Equal(calibration.Width, StyledWordBox(rendered).Width, 6);
        Assert.Equal(calibration.Document.Text, rendered.Text);
        AssertStyledWordRun(calibration.Left, left);
        AssertStyledWordRun(calibration.Right, right);

        HtmlRenderDocument intrinsic = RenderStyledWord(StyledWordMarkup(script), "display:inline-block;width:min-content");
        Assert.Equal(StyledWordAdvance(calibration.Left) + StyledWordAdvance(calibration.Right), StyledWordBox(intrinsic).Width, 6);
        Assert.Equal(calibration.Right.Y - calibration.Left.Y,
            StyledWordText(intrinsic, right.Text).Y - StyledWordText(intrinsic, left.Text).Y, 6);
        Assert.Equal(calibration.Document.Text, intrinsic.Text);
        rendered.RequireNoLoss();
        intrinsic.RequireNoLoss();
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void HtmlStyledWord_PrefixMovesTheCompleteJoinedWordToTheNextLine(bool script, bool floated) {
        const double floatWidth = 32D;
        string content = "<span id='prefix'>A </span><a href='https://example.com/word' style='color:inherit;text-decoration:none'>"
            + StyledWordMarkup(script) + "</a>";
        HtmlRenderDocument control = RenderStyledWord(content, "width:220px;white-space:nowrap");
        HtmlRenderText controlPrefix = StyledWordTexts(control).Single(text => text.Text.Trim() == "A");
        HtmlRenderText controlLeft = StyledWordText(control, script ? "steradian" : "stera");
        HtmlRenderText controlRight = StyledWordText(control, script ? "(b)" : "dian");
        double prefixWidth = controlLeft.X - controlPrefix.X;
        double leftWidth = StyledWordAdvance(controlLeft);
        double joinedWidth = leftWidth + StyledWordAdvance(controlRight);
        double lower = Math.Max(joinedWidth, prefixWidth + leftWidth);
        double width = (lower + prefixWidth + joinedWidth) / 2D;
        string measured = FormattableString.Invariant(
            $"Public calibration: P={prefixWidth:R}, L={leftWidth:R}, G={joinedWidth:R}, W={width:R}; script={script}, floated={floated}.");
        Assert.True(prefixWidth > 0D && lower + 0.001D < width && width + 0.001D < prefixWidth + joinedWidth, measured);

        string obstruction = floated
            ? "<span id='obstruction' style='float:left;width:32px;height:96px;background:gray'></span>" : string.Empty;
        HtmlRenderDocument rendered = RenderStyledWord(obstruction + content, StyledWordWidth(width + (floated ? floatWidth : 0D)));
        HtmlRenderText prefix = StyledWordTexts(rendered).Single(text => text.Text.Trim() == "A");
        HtmlRenderText left = StyledWordText(rendered, controlLeft.Text);
        HtmlRenderText right = StyledWordText(rendered, controlRight.Text);
        double origin = StyledWordBox(rendered).X;
        if (floated) {
            HtmlRenderShape actualObstruction = Assert.Single(StyledWordVisuals(rendered).OfType<HtmlRenderShape>(),
                shape => shape.Source == "span#obstruction");
            Assert.Equal(floatWidth, actualObstruction.Width, 6);
            origin = actualObstruction.X + actualObstruction.Width;
        }
        measured += FormattableString.Invariant($" Actual fresh-line origin={origin:R}; prefixY={prefix.Y:R}, leftY={left.Y:R}, rightY={right.Y:R}.");

        // The first styled part fits the old remainder; only the complete word requires relocation.
        Assert.True(prefixWidth + leftWidth < width && joinedWidth <= width, measured);
        Assert.True(left.Y > prefix.Y + 0.001D, measured);
        Assert.True(Math.Abs(right.Y - left.Y - (controlRight.Y - controlLeft.Y)) < 0.000001D, measured);
        Assert.True(Math.Abs(left.X - origin) < 0.000001D, measured);
        Assert.True(Math.Abs(right.X - left.X - leftWidth) < 0.000001D, measured);
        AssertStyledWordRun(controlLeft, left);
        AssertStyledWordRun(controlRight, right);
        Assert.Equal("https://example.com/word", left.LinkUri);
        Assert.Equal(left.LinkUri, right.LinkUri);
        Assert.Equal("A" + controlLeft.Text + controlRight.Text, StyledWordLogicalCharacters(rendered));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlStyledWord_FirstLineStyleCoversTheWholeUnbreakableWord() {
        const string content = "<span id='left'>stera</span><span id='right'>dian</span>";
        const string firstLine = "#probe::first-line{color:red}";
        HtmlRenderDocument control = RenderStyledWord(content, "width:220px;white-space:nowrap", firstLine);
        HtmlRenderText controlLeft = StyledWordText(control, "stera");
        HtmlRenderText controlRight = StyledWordText(control, "dian");
        double leftWidth = StyledWordAdvance(controlLeft), rightWidth = StyledWordAdvance(controlRight);
        double width = Math.Max(leftWidth, rightWidth) + Math.Min(leftWidth, rightWidth) / 2D;
        HtmlRenderDocument rendered = RenderStyledWord(content, StyledWordWidth(width), firstLine, "<div>After</div>");
        HtmlRenderText left = StyledWordText(rendered, "stera"), right = StyledWordText(rendered, "dian");

        Assert.Equal(controlRight.Y - controlLeft.Y, right.Y - left.Y, 6);
        Assert.Equal(left.X + leftWidth, right.X, 6);
        Assert.Equal(OfficeColor.Red, left.Color);
        Assert.Equal(OfficeColor.Red, right.Color);
        Assert.Equal(OfficeColor.Black, StyledWordText(rendered, "After").Color);
        AssertStyledWordRun(controlLeft, left);
        AssertStyledWordRun(controlRight, right);
        Assert.Equal("steradianAfter", StyledWordLogicalCharacters(rendered));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlStyledWord_FirstLinePrefixLeavesTheMovedWordInOrdinaryStyles() {
        const string content = "<span id='prefix'>A </span><span id='left'>stera</span><span id='right'>dian</span>";
        const string firstLine = "#probe::first-line{color:red}";
        HtmlRenderDocument control = RenderStyledWord(content, "width:220px;white-space:nowrap", firstLine);
        HtmlRenderText controlPrefix = StyledWordTexts(control).Single(text => text.Text.Trim() == "A");
        HtmlRenderText controlLeft = StyledWordText(control, "stera"), controlRight = StyledWordText(control, "dian");
        double prefixWidth = controlLeft.X - controlPrefix.X;
        double leftWidth = StyledWordAdvance(controlLeft), joinedWidth = leftWidth + StyledWordAdvance(controlRight);
        double lower = Math.Max(joinedWidth, prefixWidth + leftWidth);
        double width = (lower + prefixWidth + joinedWidth) / 2D;
        string measured = FormattableString.Invariant($"Public first-line prefix calibration: P={prefixWidth:R}, L={leftWidth:R}, G={joinedWidth:R}, W={width:R}.");
        Assert.True(lower + 0.001D < width && width + 0.001D < prefixWidth + joinedWidth, measured);
        HtmlRenderDocument rendered = RenderStyledWord(content, StyledWordWidth(width), firstLine);
        HtmlRenderText prefix = StyledWordTexts(rendered).Single(text => text.Text.Trim() == "A");
        HtmlRenderText left = StyledWordText(rendered, "stera"), right = StyledWordText(rendered, "dian");

        Assert.True(left.Y > prefix.Y + 0.001D, measured);
        Assert.Equal(controlRight.Y - controlLeft.Y, right.Y - left.Y, 6);
        Assert.Equal(StyledWordBox(rendered).X, left.X, 6);
        Assert.Equal(left.X + leftWidth, right.X, 6);
        Assert.Equal(OfficeColor.Red, prefix.Color);
        Assert.Equal(OfficeColor.Black, left.Color);
        Assert.Equal(OfficeColor.Black, right.Color);
        HtmlRenderDocument ordinary = RenderStyledWord(content, "width:220px;white-space:nowrap");
        AssertStyledWordRun(StyledWordText(ordinary, "stera"), left);
        AssertStyledWordRun(StyledWordText(ordinary, "dian"), right);
        Assert.Equal("Asteradian", StyledWordLogicalCharacters(rendered));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(" ")]
    [InlineData("<br>")]
    [InlineData("<wbr>")]
    public void HtmlStyledWord_AuthoredSeparatorsStillPermitTheScriptToWrap(string separator) {
        var calibration = CalibrateStyledWord(script: true);
        string content = "<span id='left'>steradian</span>" + separator + "<sup><em id='right'>(b)</em></sup>";
        HtmlRenderDocument rendered = RenderStyledWord(content, StyledWordWidth(calibration.Width));
        HtmlRenderText left = StyledWordText(rendered, "steradian"), right = StyledWordText(rendered, "(b)");

        Assert.Equal(calibration.Right.Y - calibration.Left.Y + 24D, right.Y - left.Y, 6);
        Assert.Equal(left.X, right.X, 6);
        AssertStyledWordRun(calibration.Left, left);
        AssertStyledWordRun(calibration.Right, right);
        Assert.Equal("steradian(b)", StyledWordLogicalCharacters(rendered));
        if (separator == " ") {
            HtmlRenderDocument intrinsic = RenderStyledWord(content, "display:inline-block;width:min-content");
            Assert.Equal(Math.Max(StyledWordAdvance(calibration.Left), StyledWordAdvance(calibration.Right)), StyledWordBox(intrinsic).Width, 6);
            Assert.Equal(right.Y - left.Y,
                StyledWordText(intrinsic, "(b)").Y - StyledWordText(intrinsic, "steradian").Y, 6);
            Assert.Equal(StyledWordBox(rendered).Height, StyledWordBox(intrinsic).Height, 6);
        }
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("overflow-wrap:anywhere", false)]
    [InlineData("word-break:break-all", false)]
    [InlineData("overflow-wrap:anywhere", true)]
    [InlineData("word-break:break-all", true)]
    public void HtmlStyledWord_AuthorEmergencyWrappingStillBreaksAcrossStyledFragments(string declaration, bool floated) {
        var calibration = CalibrateStyledWord(script: false);
        const double floatWidth = 32D;
        string obstruction = floated
            ? "<span id='obstruction' style='float:left;width:32px;height:96px;background:gray'></span>" : string.Empty;
        string content = obstruction + StyledWordMarkup(script: false);
        string sizing = StyledWordWidth(calibration.Width + (floated ? floatWidth : 0D)) + ";" + declaration;
        HtmlRenderDocument rendered = RenderStyledWord(content, sizing);
        HtmlRenderText[] text = StyledWordTexts(rendered).Where(item => item.Text.Length > 0).ToArray();
        double origin = 0D;
        if (floated) {
            HtmlRenderShape actualObstruction = Assert.Single(StyledWordVisuals(rendered).OfType<HtmlRenderShape>(),
                shape => shape.Source == "span#obstruction");
            Assert.Equal(floatWidth, actualObstruction.Width, 6);
            origin = actualObstruction.X + actualObstruction.Width;
            Assert.True(Math.Abs(origin - text[0].X) < 0.000001D,
                FormattableString.Invariant($"Float origin expected={origin:R}, actual={text[0].X:R}; declaration={declaration}."));
            Assert.True(text[0].Y < actualObstruction.Y + actualObstruction.Height);
        }

        Assert.True(text.Select(item => item.Y).Distinct().Count() > 1,
            declaration + "; floated=" + floated + "; geometry=" + string.Join(" | ", text.Select(item =>
                FormattableString.Invariant($"{item.Text}@({item.X:R},{item.Y:R})+{StyledWordAdvance(item):R}"))));
        Assert.Equal("steradian", string.Concat(text.Select(item => item.Text)));
        Assert.All(text, item => Assert.True(item.X >= origin - 0.001D
            && item.X + StyledWordAdvance(item) <= origin + calibration.Width + 0.001D));
        Assert.All(text.Where(item => item.Color == OfficeColor.Blue), item => Assert.Equal(calibration.Right.FontDescriptor, item.FontDescriptor));
        Assert.Equal(OfficeColor.Blue, text[text.Length - 1].Color);
        Assert.Equal("steradian", StyledWordLogicalCharacters(rendered));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlStyledWord_OverflowWrapCascadeRetainsAnywhereWithoutReplacingEarlierImportantOrValidValues() {
        const string html = "<style>#valid{overflow-wrap:normal;overflow-wrap:anywhere}"
            + "#invalid{overflow-wrap:anywhere;overflow-wrap:banana}"
            + "#important{overflow-wrap:anywhere!important;overflow-wrap:normal}"
            + "#later{overflow-wrap:anywhere;overflow-wrap:normal}"
            + "#specific{overflow-wrap:anywhere}.low{overflow-wrap:normal}</style>"
            + "<div id='valid'></div><div id='invalid'></div><div id='important' style='overflow-wrap:normal'></div>"
            + "<div id='later'></div><div id='specific' class='low'></div>"
            + "<div id='inline' style='overflow-wrap:normal;overflow-wrap:anywhere;overflow-wrap:banana'></div>";
        var document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> computed = HtmlComputedStyleEngine.Compute(document);

        foreach (string id in new[] { "valid", "invalid", "important", "specific", "inline" }) {
            Assert.Equal("anywhere", computed[document.QuerySelector("#" + id)!].GetValue("overflow-wrap"));
        }
        Assert.Equal("normal", computed[document.QuerySelector("#later")!].GetValue("overflow-wrap"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(overflow-wrap:anywhere)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(overflow-wrap:banana)"));
    }

    [Fact]
    public void HtmlStyledWord_NoWrapStillOverridesAuthorEmergencyWrapping() {
        var calibration = CalibrateStyledWord(script: true);
        HtmlRenderDocument rendered = RenderStyledWord(StyledWordMarkup(script: true),
            StyledWordWidth(calibration.Width) + ";white-space:nowrap;overflow-wrap:anywhere;word-break:break-all");
        HtmlRenderText left = StyledWordText(rendered, "steradian"), right = StyledWordText(rendered, "(b)");

        Assert.Equal(calibration.Right.Y - calibration.Left.Y, right.Y - left.Y, 6);
        Assert.Equal(left.X + StyledWordAdvance(left), right.X, 6);
        AssertStyledWordRun(calibration.Left, left);
        AssertStyledWordRun(calibration.Right, right);
        Assert.Equal(calibration.Document.Text, rendered.Text);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlStyledWord_PreferredBreakMovesOnlyTheJoinedSuffix() {
        HtmlRenderDocument control = RenderStyledWord(
            "<span>AA-</span><span>BB</span><span style='color:blue'>CC</span>", "width:220px;white-space:nowrap");
        HtmlRenderText head = StyledWordText(control, "AA-"), left = StyledWordText(control, "BB"), right = StyledWordText(control, "CC");
        double headWidth = StyledWordAdvance(head), leftWidth = StyledWordAdvance(left);
        double suffixWidth = leftWidth + StyledWordAdvance(right);
        double lower = Math.Max(suffixWidth, headWidth + leftWidth);
        double width = (lower + headWidth + suffixWidth) / 2D;
        string measured = FormattableString.Invariant($"Public preferred-break calibration: H={headWidth:R}, L={leftWidth:R}, S={suffixWidth:R}, W={width:R}.");
        Assert.True(lower + 0.001D < width && width + 0.001D < headWidth + suffixWidth, measured);
        HtmlRenderDocument rendered = RenderStyledWord("<span>AA-BB</span><span style='color:blue'>CC</span>", StyledWordWidth(width));
        HtmlRenderText[] text = StyledWordTexts(rendered).ToArray();
        string[] lines = text.GroupBy(item => item.Y).OrderBy(group => group.Key)
            .Select(group => string.Concat(group.OrderBy(item => item.X).Select(item => item.Text))).ToArray();

        Assert.True(lines.SequenceEqual(new[] { "AA-", "BBCC" }), measured + " Actual lines: " + string.Join(" | ", lines));
        HtmlRenderText suffixLeft = StyledWordText(rendered, "BB"), suffixRight = StyledWordText(rendered, "CC");
        Assert.Equal(suffixLeft.X + leftWidth, suffixRight.X, 6);
        AssertStyledWordRun(left, suffixLeft);
        AssertStyledWordRun(right, suffixRight);
        Assert.Equal("AA-BBCC", StyledWordLogicalCharacters(rendered));
        rendered.RequireNoLoss();
    }

    private static (HtmlRenderDocument Document, HtmlRenderText Left, HtmlRenderText Right, double Width) CalibrateStyledWord(bool script) {
        HtmlRenderDocument document = RenderStyledWord(StyledWordMarkup(script), "width:220px;white-space:nowrap");
        HtmlRenderText left = StyledWordText(document, script ? "steradian" : "stera");
        HtmlRenderText right = StyledWordText(document, script ? "(b)" : "dian");
        double a = StyledWordAdvance(left), b = StyledWordAdvance(right);
        double width = Math.Max(a, b) + Math.Min(a, b) / 2D;
        Assert.True(width > a + 0.001D && width > b + 0.001D && width + 0.001D < a + b);
        return (document, left, right, width);
    }

    private static HtmlRenderDocument RenderStyledWord(string content, string sizing, string extraCss = "", string followingContent = "") {
        return HtmlRenderEngine.Render(HtmlConversionDocument.Parse(StyledWordSource(content, sizing, extraCss, followingContent)), StyledWordOptions());
    }

    private static HtmlRenderOptions StyledWordOptions() {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 256D, ViewportHeight = 160D,
            Margins = HtmlRenderMargins.All(0D), DefaultFontFamily = "Styled Word Fixture", AllowSystemFontFallback = false
        };
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf"));
        // The same supplied bytes isolate fitting and carrier geometry, not independent italic-font fidelity.
        options.Fonts.Add("Styled Word Fixture", font, OfficeFontStyle.Regular);
        options.Fonts.Add("Styled Word Fixture", font, OfficeFontStyle.Italic);
        return options;
    }

    private static string StyledWordSource(string content, string sizing, string extraCss = "", string followingContent = "") =>
        "<!doctype html><html lang='en'><style>html,body{margin:0;padding:0;font:16px/24px 'Styled Word Fixture';"
            + "font-kerning:none;letter-spacing:0;word-spacing:0}#probe{margin:0;padding:0;border:0;overflow:visible;background:#eee;"
            + "white-space:normal;word-break:normal;overflow-wrap:normal;line-break:auto;hyphens:none;" + sizing + "}"
            + "sup{font-size:75%;vertical-align:super}em{font-style:italic}" + extraCss
            + "</style><body><div id='probe'>" + content + "</div>" + followingContent + "</body></html>";

    private static string StyledWordMarkup(bool script) => script
        ? "<span id='left'>steradian</span><sup><em id='right'>(b)</em></sup>"
        : "<span id='left'>stera</span><span id='right' style='color:blue'>dian</span>";

    private static string StyledWordWidth(double width) => "width:" + width.ToString("R", CultureInfo.InvariantCulture) + "px";
    private static IEnumerable<HtmlRenderVisual> StyledWordVisuals(HtmlRenderDocument document) => document.Pages.SelectMany(page => page.Visuals);
    private static IEnumerable<HtmlRenderText> StyledWordTexts(HtmlRenderDocument document) => StyledWordVisuals(document).OfType<HtmlRenderText>();
    private static HtmlRenderText StyledWordText(HtmlRenderDocument document, string value) => Assert.Single(StyledWordTexts(document), text => text.Text == value);
    private static HtmlRenderShape StyledWordBox(HtmlRenderDocument document) => Assert.Single(StyledWordVisuals(document).OfType<HtmlRenderShape>(), shape => shape.Source == "div#probe");
    private static double StyledWordAdvance(HtmlRenderText text) => Assert.IsType<double>(text.TextAdvanceWidth);

    // Public searchable text inserts separators between visuals; source characters must still occur exactly once in source order.
    private static string StyledWordLogicalCharacters(HtmlRenderDocument document) => new(document.Text.Where(character => !char.IsWhiteSpace(character)).ToArray());

    private static void AssertStyledWordRun(HtmlRenderText expected, HtmlRenderText actual) {
        Assert.Equal(expected.Font.FamilyName, actual.Font.FamilyName);
        Assert.Equal(expected.Font.Size, actual.Font.Size, 6);
        Assert.Equal(expected.Font.Style, actual.Font.Style);
        Assert.Equal(expected.FontDescriptor, actual.FontDescriptor);
        Assert.Equal(expected.Baseline, actual.Baseline);
        Assert.Equal(expected.BaselineLevel, actual.BaselineLevel);
        Assert.Equal(expected.BaselineScale, actual.BaselineScale, 6);
        Assert.Equal(expected.BaselineOffset, actual.BaselineOffset, 6);
        Assert.Equal(expected.Color, actual.Color);
        Assert.Equal(expected.LinkUri, actual.LinkUri);
        Assert.Equal(expected.Source, actual.Source);
        Assert.Equal(StyledWordAdvance(expected), StyledWordAdvance(actual), 6);
        Assert.Equal(expected.TextPaintWidth, actual.TextPaintWidth);
    }
}

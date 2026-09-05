using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("margin-left:5px;margin-inline-start:20px", 20D)]
    [InlineData("margin-inline-start:20px;margin-left:5px", 5D)]
    [InlineData("margin-inline-start:20px!important;margin-left:5px", 20D)]
    [InlineData("margin-inline-start:20px;margin:5px", 5D)]
    [InlineData("margin:5px;margin-inline-start:20px", 20D)]
    [InlineData("margin-left:20px;margin:5px", 5D)]
    [InlineData("margin:5px;margin-left:20px", 20D)]
    [InlineData("margin-inline-start:20px;margin:5px!important", 5D)]
    [InlineData("margin:5px!important;margin-inline-start:20px", 5D)]
    [InlineData("--gap:5px;margin-inline-start:20px;margin:var(--gap)", 5D)]
    public void HtmlRender_LogicalAndPhysicalPropertiesRespectTheFullCascade(string declarations, double expectedX) {
        string html = $"<div id='logical' style='{declarations};width:20px;height:10px;background:red'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape shape = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#logical" && item.Shape.FillColor == OfficeColor.Red);

        Assert.Equal(expectedX, shape.X, 3);
    }

    [Theory]
    [InlineData("#logical { margin:5px } .box { margin-inline-start:20px }", 5D)]
    [InlineData("#logical { margin-inline-start:20px } .box { margin:5px }", 20D)]
    public void HtmlRender_ShorthandProjectionPreservesSelectorSpecificity(string css, double expectedX) {
        string html = $"<style>{css}</style><div id='logical' class='box' style='width:20px;height:10px;background:red'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape shape = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#logical" && item.Shape.FillColor == OfficeColor.Red);

        Assert.Equal(expectedX, shape.X, 3);
    }

    [Theory]
    [InlineData("padding-inline-start:20px;padding:5px", 5D)]
    [InlineData("padding:5px;padding-inline-start:20px", 20D)]
    [InlineData("padding-left:20px;padding:5px", 5D)]
    [InlineData("padding:5px;padding-left:20px", 20D)]
    [InlineData("padding-inline-start:20px;padding:5px!important", 5D)]
    public void HtmlRender_PaddingShorthandsAndLogicalLonghandsRespectTheFullCascade(string declarations, double expectedX) {
        string html = $"<div id='parent' style='{declarations};width:60px'><div id='child' style='width:10px;height:10px;background:red'></div></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape shape = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#child" && item.Shape.FillColor == OfficeColor.Red);

        Assert.Equal(expectedX, shape.X, 3);
    }

    [Theory]
    [InlineData("border-inline-start:5px solid red;border:2px solid blue", 2D, 0, 0, 255)]
    [InlineData("border:2px solid blue;border-inline-start:5px solid red", 5D, 255, 0, 0)]
    [InlineData("border-left-width:8px;border:2px solid blue", 2D, 0, 0, 255)]
    [InlineData("border:2px solid blue;border-left-width:8px", 8D, 0, 0, 255)]
    [InlineData("border-inline-start:5px solid red;border:2px solid blue!important", 2D, 0, 0, 255)]
    public void HtmlRender_BorderShorthandsComponentsAndLogicalSidesRespectTheFullCascade(
        string declarations,
        double expectedWidth,
        byte red,
        byte green,
        byte blue) {
        string html = $"<div id='logical' style='{declarations};width:30px;height:20px'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        IReadOnlyList<HtmlRenderShape> strokes = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(item => item.Shape.StrokeWidth > 0D)
            .ToList();
        HtmlRenderShape? border = strokes.FirstOrDefault(item =>
            item.Source.EndsWith(":border-left", StringComparison.Ordinal));
        border ??= Assert.Single(strokes);

        Assert.Equal(expectedWidth, border.Shape.StrokeWidth, 3);
        Assert.Equal(OfficeColor.FromRgb(red, green, blue), border.Shape.StrokeColor);
    }

    [Theory]
    [InlineData("ltr", "div#logical:border-bottom")]
    [InlineData("rtl", "div#logical:border-top")]
    public void HtmlRender_SidewaysLrUsesBottomToTopInlineDirection(string direction, string expectedBorderSource) {
        string html = $"<div id='logical' dir='{direction}' style='writing-mode:sideways-lr;"
            + "border-inline-start:3px solid blue;width:30px;height:50px;background:red'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == expectedBorderSource && item.Shape.StrokeColor == OfficeColor.Blue);
    }

    [Fact]
    public void HtmlRender_LogicalBoxPropertiesFollowVerticalWritingMode() {
        const string html = "<div id='logical' style='writing-mode:vertical-rl;inline-size:80px;block-size:40px;"
            + "padding-inline-start:6px;border-block-start:3px solid #123456;box-sizing:border-box;"
            + "background:#abcdef;font-size:8px;line-height:10px'>Logical</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 200D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        IReadOnlyList<HtmlRenderVisual> visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToList();
        HtmlRenderShape background = Assert.Single(visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#logical" && shape.Shape.FillColor == OfficeColor.FromRgb(0xAB, 0xCD, 0xEF));
        HtmlRenderText text = Assert.Single(visuals.OfType<HtmlRenderText>(), item => item.Text == "L");

        Assert.Equal(40D, background.Width, 3);
        Assert.Equal(80D, background.Height, 3);
        Assert.Equal(6D, text.Y + text.Height / 2D - (text.TextAdvanceWidth ?? text.Width) / 2D, 3);
        Assert.Contains(visuals.OfType<HtmlRenderShape>(), shape =>
            shape.Source == "div#logical:border-right"
            && shape.Shape.StrokeWidth == 3D
            && shape.Shape.StrokeColor == OfficeColor.FromRgb(0x12, 0x34, 0x56));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(writing-mode:vertical-rl)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-block-start:3px solid #123456)"));
    }

    [Fact]
    public void HtmlRender_LogicalInsetsAndSizesPositionVerticalAbsoluteBoxes() {
        const string html = "<div style='position:relative;width:120px;height:100px;margin:0'>"
            + "<div id='positioned' style='position:absolute;writing-mode:vertical-rl;"
            + "inset-inline-start:10px;inset-block-start:15px;inline-size:20px;block-size:30px;"
            + "background:#ff0000'></div></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 160D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape positioned = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#positioned" && shape.Shape.FillColor == OfficeColor.Red);

        Assert.Equal(75D, positioned.X, 3);
        Assert.Equal(10D, positioned.Y, 3);
        Assert.Equal(30D, positioned.Width, 3);
        Assert.Equal(20D, positioned.Height, 3);
    }

    [Theory]
    [InlineData("sideways-rl", 0D, 1D, -1D, 0D)]
    [InlineData("sideways-lr", 0D, -1D, 1D, 0D)]
    public void HtmlRender_SidewaysWritingModesUseSearchableAffineText(
        string writingMode,
        double m11,
        double m12,
        double m21,
        double m22) {
        string html = $"<div id='vertical' style='writing-mode:{writingMode};inline-size:90px;block-size:30px;"
            + "margin:0;background:#eeeeee;font-size:10px;line-height:12px'>Vertical text</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 140D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderEffectGroup vertical = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderEffectGroup>(),
            group => group.Source == "div#vertical:vertical-writing");
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(140D / HtmlRenderOptions.CssPixelsPerInch, 120D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        Assert.Equal(m11, vertical.Transform.M11, 6);
        Assert.Equal(m12, vertical.Transform.M12, 6);
        Assert.Equal(m21, vertical.Transform.M21, 6);
        Assert.Equal(m22, vertical.Transform.M22, 6);
        Assert.Contains("Vertical text", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_VerticalMixedOrientationKeepsCjkUprightAndLatinSideways() {
        string html = "<style>" + CreatePortableEmbeddedFontFaceCss("Vertical Test", 0x7E26, 0x66F8, 0x304D)
            + "</style><div style=\"font-family:'Vertical Test';writing-mode:vertical-rl;inline-size:80px;block-size:30px\">縦書き Latin</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        IReadOnlyList<HtmlRenderVisual> visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToList();
        IReadOnlyList<HtmlRenderText> text = visuals.OfType<HtmlRenderText>().ToList();
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions {
            PageSize = new OfficePageSize(160D / HtmlRenderOptions.CssPixelsPerInch, 120D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains(text, glyph => glyph.Text == "縦");
        Assert.Contains(text, glyph => glyph.Text == "書");
        Assert.Contains(visuals.OfType<HtmlRenderEffectGroup>(), group =>
            group.Visuals.OfType<HtmlRenderText>().Any(glyph => glyph.Text == "L")
            && group.Transform.M12 == 1D);
        Assert.Contains("縦書きLatin", PdfCore.PdfReadDocument.Open(pdf).ExtractText().Replace(" ", string.Empty), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("over", true)]
    [InlineData("under", false)]
    public void HtmlRender_RubyPositionsScaledAnnotationsAroundTheBase(string position, bool annotationAbove) {
        string html = $"<p style='font-size:20px;line-height:24px;margin:0'><ruby style='ruby-position:{position};ruby-align:start'>"
            + "<rb>東</rb><rt>とう</rt></ruby></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        IReadOnlyList<HtmlRenderText> text = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .ToList();
        HtmlRenderText rubyBase = Assert.Single(text, item => item.Text == "東");
        HtmlRenderText annotation = Assert.Single(text, item => item.Text == "とう");

        Assert.Equal(20D, rubyBase.Font.Size, 3);
        Assert.Equal(10D, annotation.Font.Size, 3);
        Assert.Equal(annotationAbove, annotation.Y < rubyBase.Y);
        Assert.Equal(rubyBase.X, annotation.X, 3);
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(ruby-position:under)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(ruby-align:start)"));
    }

    [Fact]
    public void HtmlRender_VerticalRubyPlacesOverAnnotationOnTheLineOuterSide() {
        const string html = "<div style='writing-mode:vertical-rl;inline-size:80px;block-size:40px;font-size:20px;line-height:24px'>"
            + "<ruby><rb>東</rb><rt>とう</rt></ruby></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        IReadOnlyList<HtmlRenderText> text = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .ToList();
        HtmlRenderText rubyBase = Assert.Single(text, item => item.Text == "東");
        IReadOnlyList<HtmlRenderText> annotations = text.Where(item => item.Text == "と" || item.Text == "う").ToList();

        Assert.Equal(2, annotations.Count);
        Assert.All(annotations, annotation => Assert.True(annotation.X > rubyBase.X));
    }

    [Theory]
    [InlineData("vertical-rl", true)]
    [InlineData("vertical-lr", false)]
    public void HtmlRender_VerticalBlockChildrenAdvanceInTheBlockDirection(string writingMode, bool firstIsRightmost) {
        string html = $"<div style='writing-mode:{writingMode};inline-size:80px;block-size:80px;margin:0;font-size:12px;line-height:16px'>"
            + "<p style='margin:0'>First</p><p style='margin:0'>Second</p></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        IReadOnlyList<HtmlRenderText> text = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .ToList();
        double firstX = Assert.Single(text, item => item.Text == "F").X;
        double secondX = Assert.Single(text, item => item.Text == "S").X;
        Assert.Equal(firstIsRightmost, firstX > secondX);
    }

    [Theory]
    [InlineData("vertical-rl", true)]
    [InlineData("vertical-lr", false)]
    public void HtmlRender_VerticalNestedAndDecoratedBlocksAdvanceByTheirPaintedExtent(string writingMode, bool followingIsLeft) {
        string html = $"<div style='writing-mode:{writingMode};inline-size:90px;block-size:180px;margin:0;font-size:12px;line-height:16px'>"
            + "<div id='nested' style='width:56px;margin:0 8px 0 0;padding:4px;border:2px solid #123;background:#def'>"
            + "<span id='nested-a' style='display:block'>A</span><span id='nested-b' style='display:block'>B</span></div>"
            + "<p id='following' style='margin:0'>C</p></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        IReadOnlyList<HtmlRenderText> text = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .ToList();
        double[] nestedX = text
            .Where(item => item.Text == "A" || item.Text == "B")
            .Select(item => item.X)
            .ToArray();
        double followingX = Assert.Single(text, item => item.Text == "C").X;

        Assert.Equal(2, nestedX.Length);
        if (followingIsLeft) Assert.True(followingX < nestedX.Min() - 20D);
        else Assert.True(followingX > nestedX.Max() + 20D);
    }

    [Fact]
    public void HtmlRender_FirstLetterAndFirstLineStyleTheRenderedFragments() {
        const string html = "<style>#lead::first-letter{font-size:30px;color:#cc0000}"
            + "#lead::first-line{font-weight:bold;color:#0000cc}</style>"
            + "<p id='lead' style='width:90px;margin:0;font-size:12px;line-height:16px'>Hello world across two lines</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        IReadOnlyList<HtmlRenderText> text = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .ToList();
        HtmlRenderText firstLetter = Assert.Single(text, item => item.Text == "H");
        double firstLineY = text.Min(item => item.Y);

        Assert.Equal(30D, firstLetter.Font.Size, 3);
        Assert.Equal(OfficeColor.FromRgb(0xCC, 0x00, 0x00), firstLetter.Color);
        Assert.Contains(text, item => item.Y <= firstLineY + 0.001D
            && item.Text.Contains("ello", StringComparison.Ordinal)
            && item.Color == OfficeColor.FromRgb(0x00, 0x00, 0xCC)
            && item.Font.IsBold);
        Assert.Contains(text, item => item.Y > firstLineY + 0.001D
            && item.Color == OfficeColor.Black);
    }

    [Fact]
    public void HtmlRender_FirstLetterIncludesLeadingPunctuation() {
        const string html = "<style>p::first-letter{color:red}</style><p style='margin:0'>“A” sample</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderText firstLetter = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            item => item.Text == "“A”");

        Assert.Equal(OfficeColor.Red, firstLetter.Color);
    }

    [Fact]
    public void HtmlRender_FirstLineStopsAtAnEmergencyBreakInsideOneToken() {
        const string html = "<style>#lead::first-line{color:#0000cc}</style>"
            + "<p id='lead' style='width:48px;margin:0;font-size:12px;line-height:16px;overflow-wrap:anywhere'>"
            + "Supercalifragilisticexpialidocious</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        IReadOnlyList<HtmlRenderText> text = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .ToList();
        double firstLineY = text.Min(item => item.Y);

        Assert.Contains(text, item => item.Y <= firstLineY + 0.001D
            && item.Color == OfficeColor.FromRgb(0x00, 0x00, 0xCC));
        Assert.Contains(text, item => item.Y > firstLineY + 0.001D
            && item.Color == OfficeColor.Black);
        Assert.DoesNotContain(text, item => item.Y > firstLineY + 0.001D
            && item.Color == OfficeColor.FromRgb(0x00, 0x00, 0xCC));
    }

    [Fact]
    public void HtmlRender_FirstLineStylesAnEmergencyTokenPrefixAfterEarlierContent() {
        const string html = "<style>#lead::first-line{color:#0000cc}</style>"
            + "<p id='lead' style='width:72px;margin:0;font-size:12px;line-height:16px;overflow-wrap:anywhere'>"
            + "Hi supercalifragilisticexpialidocious</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        IReadOnlyList<HtmlRenderText> text = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .ToList();
        double firstLineY = text.Min(item => item.Y);

        Assert.Contains(text, item => item.Y <= firstLineY + 0.001D && item.Text.Contains("Hi", StringComparison.Ordinal));
        Assert.DoesNotContain(text, item => item.Y <= firstLineY + 0.001D && item.Color == OfficeColor.Black);
        Assert.Contains(text, item => item.Y > firstLineY + 0.001D && item.Color == OfficeColor.Black);
        Assert.DoesNotContain(text, item => item.Y > firstLineY + 0.001D
            && item.Color == OfficeColor.FromRgb(0x00, 0x00, 0xCC));
    }
}

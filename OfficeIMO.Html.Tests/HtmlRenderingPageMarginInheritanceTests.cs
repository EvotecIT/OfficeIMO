using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using System.Threading.Tasks;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlPageMarginInheritance_PublicRenderRoutesInheritTheRootRatherThanTheBody(bool asynchronous) {
        const string html = """
            <style>
              html { font:italic bold 20px/1.5 Pinned; color:red; }
              body,p { margin:0; font:9px/12px Arial; color:blue; }
              @page { size:300px 240px; margin:60px;
                @bottom-center { content:'Margin'; }
              }
            </style><p>BodyOnly</p>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderOptions options = PageMarginTypographyOptions();

        HtmlRenderDocument rendered = asynchronous
            ? await HtmlRenderEngine.RenderAsync(document, options)
            : HtmlRenderEngine.Render(document, options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText margin = PageMarginTypographyText(page);
        HtmlRenderText body = Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>(),
            text => text.Text == "BodyOnly");

        Assert.Equal("Pinned", margin.Font.FamilyName);
        Assert.Equal(20D, margin.Font.Size, 6);
        Assert.Equal(OfficeFontStyle.Bold | OfficeFontStyle.Italic, margin.Font.Style);
        Assert.Equal(30D, margin.LineHeight, 6);
        Assert.Equal(OfficeColor.Red, margin.Color);
        Assert.Equal("Arial", body.Font.FamilyName);
        Assert.Equal(9D, body.Font.Size, 6);
        Assert.Equal(OfficeFontStyle.Regular, body.Font.Style);
        Assert.Equal(12D, body.LineHeight, 6);
        Assert.Equal(OfficeColor.Blue, body.Color);
    }

    [Fact]
    public void HtmlPageMarginInheritance_TypographyOnlyPageRulesStyleTheirMatchingPages() {
        const string html = """
            <style>
              html { font:16px/1.25 Arial; color:black; }
              body,p { margin:0; font:9px/12px Arial; color:blue; }
              @page { size:300px 240px; margin:60px;
                @bottom-center { content:'Margin'; }
              }
              @page :first { font:italic bold 24px/1.5 Pinned; color:red; }
              @page :left { font-size:12px; line-height:2; color:green; }
            </style><p>First</p><p style='break-before:page'>Second</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions());

        Assert.Equal(2, rendered.Pages.Count);
        HtmlRenderText first = PageMarginTypographyText(rendered.Pages[0]);
        HtmlRenderText left = PageMarginTypographyText(rendered.Pages[1]);
        Assert.Equal("Pinned", first.Font.FamilyName);
        Assert.Equal(24D, first.Font.Size, 6);
        Assert.Equal(36D, first.LineHeight, 6);
        Assert.Equal(OfficeFontStyle.Bold | OfficeFontStyle.Italic, first.Font.Style);
        Assert.Equal(OfficeColor.Red, first.Color);
        Assert.Equal("Arial", left.Font.FamilyName);
        Assert.Equal(12D, left.Font.Size, 6);
        Assert.Equal(24D, left.LineHeight, 6);
        Assert.Equal(OfficeFontStyle.Regular, left.Font.Style);
        Assert.Equal(OfficeColor.Green, left.Color);
        Assert.All(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().Where(text => text.SemanticRole != "page-margin"),
            text => Assert.Equal(OfficeColor.Blue, text.Color));
    }

    [Fact]
    public void HtmlPageMarginInheritance_ExplicitMarginValuesBeatImportantInheritedPageValues() {
        const string html = """
            <style>
              html { font:16px/1.25 Pinned; }
              body,p { margin:0; font:9px/12px Pinned; }
              @page :first { font-size:30px !important; color:green !important; line-height:2 !important; }
              @page { size:300px 240px; margin:60px;
                @bottom-center { content:'Margin'; font-size:10px; line-height:12px; color:blue; }
              }
            </style><p>BodyOnly</p>
            """;

        HtmlRenderText margin = PageMarginTypographyText(Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages));

        Assert.Equal(10D, margin.Font.Size, 6);
        Assert.Equal(12D, margin.LineHeight, 6);
        Assert.Equal(OfficeColor.Blue, margin.Color);
    }

    [Theory]
    [InlineData("inherit", 24D, 36D, true)]
    [InlineData("unset", 24D, 36D, true)]
    [InlineData("revert", 24D, 36D, true)]
    [InlineData("initial", 16D, 19D, false)]
    public void HtmlPageMarginInheritance_MarginCssWideValuesUseThePageOrPropertyInitial(
        string keyword, double expectedSize, double expectedLeading, bool inheritsStyle) {
        string declarations = "font-size:30px;font-weight:normal;font-style:normal;line-height:10px;color:blue;"
            + string.Join(";", new[] { "font-size", "font-weight", "font-style", "line-height", "color" }
                .Select(property => property + ":" + keyword));
        string html = PageMarginInheritanceSource(
            "font:20px/1.25 Pinned;color:green", "font:italic bold 24px/1.5 Pinned;color:red", declarations);

        HtmlRenderText margin = PageMarginTypographyText(Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages));

        Assert.Equal(expectedSize, margin.Font.Size, 6);
        Assert.Equal(expectedLeading, margin.LineHeight, 6);
        Assert.Equal(inheritsStyle ? OfficeFontStyle.Bold | OfficeFontStyle.Italic : OfficeFontStyle.Regular,
            margin.Font.Style);
        Assert.Equal(inheritsStyle ? OfficeColor.Red : OfficeColor.Black, margin.Color);
    }

    [Theory]
    [InlineData("inherit", 20D)]
    [InlineData("unset", 20D)]
    [InlineData("revert", 20D)]
    [InlineData("initial", 16D)]
    public void HtmlPageMarginInheritance_PageCssWideValuesResolveBeforeMarginInheritance(
        string keyword, double expectedSize) {
        string html = PageMarginInheritanceSource("font:20px/1.5 Pinned",
            "font-size:30px;font-size:" + keyword, string.Empty);

        HtmlRenderText margin = PageMarginTypographyText(Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages));

        Assert.Equal(expectedSize, margin.Font.Size, 6);
        Assert.Equal(expectedSize * 1.5D, margin.LineHeight, 6);
    }

    [Fact]
    public void HtmlPageMarginInheritance_LayerRollbackKeepsTheLowerMarginValueInsteadOfItsParent() {
        const string html = """
            <style>
              @layer base,theme;
              html { font:20px/1.25 Pinned; }
              body,p { margin:0; font:9px/12px Pinned; }
              @layer base { @page { size:300px 240px; margin:60px;
                @bottom-center { content:'Margin'; font:12px/1.5 Pinned; color:blue; }
              } }
              @layer theme { @page { font-size:24px; color:red;
                @bottom-center { font:30px/2 Other; color:green; font:revert-layer; color:revert-layer; }
              } }
            </style><p>BodyOnly</p>
            """;

        HtmlRenderText margin = PageMarginTypographyText(Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages));

        Assert.Equal("Pinned", margin.Font.FamilyName);
        Assert.Equal(12D, margin.Font.Size, 6);
        Assert.Equal(18D, margin.LineHeight, 6);
        Assert.Equal(OfficeColor.Blue, margin.Color);
    }

    [Theory]
    [InlineData("150%", "", 30D)]
    [InlineData("1.5em", "", 30D)]
    [InlineData("150%", "font-size:inherit", 30D)]
    [InlineData("150%", "font-size:50%", 15D)]
    [InlineData("150%", "font-size:0.5em", 15D)]
    [InlineData("150%", "font-size:2rem", 40D)]
    public void HtmlPageMarginInheritance_RelativeFontSizesUseTheirOwnContextOnce(
        string pageSize, string marginDeclaration, double expectedSize) {
        string html = PageMarginInheritanceSource("font:20px/1.5 Pinned",
            "font-size:" + pageSize, marginDeclaration);

        HtmlRenderText margin = PageMarginTypographyText(Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages));

        Assert.Equal(expectedSize, margin.Font.Size, 6);
        Assert.Equal(expectedSize * 1.5D, margin.LineHeight, 6);
    }

    [Theory]
    [InlineData("1.5", "", 15D)]
    [InlineData("150%", "", 30D)]
    [InlineData("1.5em", "", 30D)]
    [InlineData("30px", "", 30D)]
    [InlineData("1.5", "line-height:1.2", 12D)]
    [InlineData("1.5", "line-height:120%", 36D)]
    [InlineData("1.5", "line-height:1.2em", 36D)]
    public void HtmlPageMarginInheritance_PercentageAndLengthLeadingFreezeWhileNumbersScale(
        string rootLeading, string pageLeading, double expectedLeading) {
        string html = PageMarginInheritanceSource("font:20px/" + rootLeading + " Pinned",
            "font-size:30px;" + pageLeading, "font-size:10px");

        HtmlRenderText margin = PageMarginTypographyText(Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages));

        Assert.Equal(10D, margin.Font.Size, 6);
        Assert.Equal(expectedLeading, margin.LineHeight, 6);
    }

    [Theory]
    [InlineData("bottom-left", "", OfficeTextAlignment.Left)]
    [InlineData("bottom-center", "", OfficeTextAlignment.Center)]
    [InlineData("bottom-right", "", OfficeTextAlignment.Right)]
    [InlineData("bottom-center", "text-align:inherit", OfficeTextAlignment.Right)]
    [InlineData("bottom-center", "text-align:unset", OfficeTextAlignment.Right)]
    [InlineData("bottom-center", "text-align:initial", OfficeTextAlignment.Left)]
    public void HtmlPageMarginInheritance_OmittedAlignmentKeepsItsPositionWhileExplicitInheritanceUsesThePage(
        string position, string marginDeclaration, OfficeTextAlignment expectedAlignment) {
        string html = PageMarginInheritanceSource("font:16px/1.25 Pinned;text-align:left",
            "text-align:right", marginDeclaration, position);

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages);

        Assert.Equal(expectedAlignment, PageMarginTypographyText(page, position).Alignment);
    }

    [Fact]
    public void HtmlPageMarginInheritance_CustomPropertyTokensResolveWhereUsedAndHonorFallbacks() {
        const string html = """
            <style>
              html { font:20px/1.25 Pinned; --size:150%; --leading:150%; --family:Pinned; --ink:red; }
              body,p { margin:0; font:9px/12px Pinned; }
              @page { size:300px 240px; margin:60px;
                font-family:var(--family); font-size:var(--size); line-height:var(--leading); --ink:green;
                @bottom-left { content:'A'; color:var(--ink); }
                @bottom-center { content:'B'; font-size:var(--size); }
                @bottom-right { content:'C'; font-size:var(--missing-size,10px);
                  line-height:var(--missing-leading,2); font-family:var(--missing-family,Pinned); }
              }
            </style><p>BodyOnly</p>
            """;

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages);
        HtmlRenderText inherited = PageMarginTypographyText(page, "bottom-left");
        HtmlRenderText substituted = PageMarginTypographyText(page);
        HtmlRenderText fallback = PageMarginTypographyText(page, "bottom-right");

        Assert.Equal(30D, inherited.Font.Size, 6);
        Assert.Equal(45D, inherited.LineHeight, 6);
        Assert.Equal(OfficeColor.Green, inherited.Color);
        Assert.Equal(45D, substituted.Font.Size, 6);
        Assert.Equal(45D, substituted.LineHeight, 6);
        Assert.Equal("Pinned", fallback.Font.FamilyName);
        Assert.Equal(10D, fallback.Font.Size, 6);
        Assert.Equal(20D, fallback.LineHeight, 6);
    }

    [Theory]
    [InlineData("", 12D)]
    [InlineData("--n:30px;", 30D)]
    [InlineData("--n:red;", 12D)]
    public void HtmlPageMarginInheritance_RegisteredNonInheritedValuesResetInEachGeneratedContext(
        string pageDeclaration, double expectedPageSize) {
        string html = "<style>@property --n{syntax:'<length>';inherits:false;initial-value:12px}"
            + "html{font:20px/1.5 Pinned;--n:30px}body,p{margin:0;font:9px/12px Pinned}"
            + "@page{size:300px 240px;margin:60px;" + pageDeclaration + "font-size:var(--n);"
            + "@bottom-left{content:'A'}@bottom-center{content:'B';font-size:var(--n)}}"
            + "</style><p>BodyOnly</p>";

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages);

        Assert.Equal(expectedPageSize, PageMarginTypographyText(page, "bottom-left").Font.Size, 6);
        Assert.Equal(12D, PageMarginTypographyText(page).Font.Size, 6);
    }

    [Fact]
    public void HtmlPageMarginInheritance_GeneratedContextsShareTheRegisteredPropertyFanoutBudget() {
        string registrations = string.Concat(Enumerable.Range(0, 20).Select(index =>
            "@property --n" + index + "{syntax:'<length>';inherits:false;initial-value:12px}"));
        string[] positions = {
            "top-left-corner", "top-left", "top-center", "top-right", "top-right-corner",
            "right-top", "right-middle", "right-bottom", "bottom-right-corner",
            "bottom-right", "bottom-center", "bottom-left", "bottom-left-corner",
            "left-bottom", "left-middle", "left-top"
        };
        string marginRules = string.Concat(positions.Select(position =>
            "@" + position + "{content:'A'}"));
        string Source(string margins) => "<style>" + registrations
            + "html{font:12px/1.5 Pinned}body,p{margin:0}@page{size:300px 240px;margin:60px;"
            + margins + "}</style><p>BodyOnly</p>";
        var parsing = new HtmlConversionDocumentOptions { Trust = HtmlInputTrust.Untrusted };
        parsing.Limits.MaxCssDeclarations = 300;

        Assert.Single(HtmlRenderEngine.Render(HtmlConversionDocument.Parse(Source(string.Empty), parsing),
            PageMarginTypographyOptions()).Pages);
        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() => HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(Source(marginRules), parsing), PageMarginTypographyOptions()));

        Assert.Equal("MaxCssDeclarations", error.LimitSource);
    }

    [Fact]
    public void HtmlPageMarginInheritance_ExplicitAlignmentInheritanceWithoutParentAlignmentUsesTheInitialLeft() {
        string html = PageMarginInheritanceSource("font:16px/1.25 Pinned",
            string.Empty, "text-align:inherit");

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages);

        Assert.Equal(OfficeTextAlignment.Left, PageMarginTypographyText(page).Alignment);
    }

    [Fact]
    public void HtmlPageMarginInheritance_DeferredFontShorthandUsesFinalLonghandsAndInvalidValuesInherit() {
        const string html = """
            <style>
              html { font:20px/1.25 Pinned; --page-font:italic bold 150%/1.5 Pinned; --margin-font:10px/150% Pinned; }
              body,p { margin:0; font:9px/12px Pinned; }
              @page { size:300px 240px; margin:60px; font:var(--page-font);
                @bottom-left { content:'A'; font:var(--margin-font); font-size:12px; }
                @bottom-center { content:'B'; font-weight:normal; font:var(--missing); }
                @bottom-right { content:'C'; font:var(--missing,10px/2 Pinned); }
              }
            </style><p>BodyOnly</p>
            """;

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages);
        HtmlRenderText longhand = PageMarginTypographyText(page, "bottom-left");
        HtmlRenderText invalid = PageMarginTypographyText(page);
        HtmlRenderText fallback = PageMarginTypographyText(page, "bottom-right");

        Assert.Equal(12D, longhand.Font.Size, 6);
        Assert.Equal(18D, longhand.LineHeight, 6);
        Assert.Equal(OfficeFontStyle.Regular, longhand.Font.Style);
        Assert.Equal(30D, invalid.Font.Size, 6);
        Assert.Equal(45D, invalid.LineHeight, 6);
        Assert.Equal(OfficeFontStyle.Bold | OfficeFontStyle.Italic, invalid.Font.Style);
        Assert.Equal(10D, fallback.Font.Size, 6);
        Assert.Equal(20D, fallback.LineHeight, 6);
        Assert.Equal(OfficeFontStyle.Regular, fallback.Font.Style);
    }

    [Fact]
    public void HtmlPageMarginInheritance_SuppliedPageFaceSetsNaturalMetricsBeforeMarginPainting() {
        var options = new HtmlToPdfOptions(PageMarginTypographyOptions()) {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        options.Fonts.Add("MarginFace", ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(
            800, -200, 100, 200, 'A', 'B'));
        const string html = """
            <style>
              html { font:16px Pinned; }
              body,p { margin:0; font:9px/12px Pinned; }
              @page { size:300px 240px; margin:60px; font:20px MarginFace;
                @bottom-center { content:'A\A B'; }
              }
            </style><p>BodyOnly</p>
            """;

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), options).Pages);
        HtmlRenderText margin = PageMarginTypographyText(page);
        OfficeDrawingText[] lines = PageMarginTypographyDrawingText(page.CreateDrawing())
            .Where(text => text.Text == "A" || text.Text == "B").ToArray();

        Assert.Equal("MarginFace", margin.Font.FamilyName);
        Assert.Equal(20D, margin.Font.Size, 6);
        Assert.Equal(22D, margin.LineHeight, 6);
        Assert.Equal(2, lines.Length);
        Assert.Equal(22D, lines[1].Y - lines[0].Y, 6);
    }

    [Fact]
    public void HtmlPageMarginInheritance_LibraryFontPreparationUsesInheritedPageTypography() {
        string? installedFamily = PdfCore.PdfCreateOptions.DefaultDocumentFontFamilyFallback.Split(',')
            .Select(candidate => candidate.Trim())
            .FirstOrDefault(candidate => PdfCore.PdfEmbeddedFontFamily.TryFromSystem(candidate, out _));
        // Match the existing optional installed-font integration lane. The
        // remaining inheritance cases use pinned supplied faces on every host.
        if (installedFamily == null) return;
        var family = PdfCore.PdfEmbeddedFontFamily.FromSystem(installedFamily);
        const string header = "łłłłłłłłłłłł";
        string html = "<style>html{font:20px Arial}body,p{margin:0;font:9px/12px Arial}"
            + "@page{size:100px 240px;margin:100px 10px 20px;@top-center{content:'" + header + "'}}"
            + "@page:first{font-size:150%}</style><p>Body</p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);

        HtmlRenderPage automatic = Assert.Single(document.RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Pdf, new HtmlToPdfOptions { AllowSystemFontFallback = true })).RenderResult.Document.Pages);
        HtmlRenderPage explicitFont = Assert.Single(document.RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Pdf, new HtmlToPdfOptions { FontFamily = family, AllowSystemFontFallback = false })).RenderResult.Document.Pages);
        HtmlRenderText margin = PageMarginTypographyText(automatic, "top-center");
        HtmlRenderText control = PageMarginTypographyText(explicitFont, "top-center");

        Assert.Equal(30D, margin.Font.Size, 6);
        Assert.Equal(control.LineHeight, margin.LineHeight, 6);
        Assert.Equal(control.Height, margin.Height, 6);
        Assert.Equal(control.Y, margin.Y, 6);
        Assert.Equal(header, margin.Text);
    }

    private static string PageMarginInheritanceSource(string rootDeclarations,
        string pageDeclarations, string marginDeclarations, string position = "bottom-center") =>
        "<!doctype html><style>html{" + rootDeclarations + "}body,p{margin:0;font:9px/12px Pinned}"
        + "@page{size:300px 240px;margin:60px;" + pageDeclarations + ";@" + position
        + "{content:'Margin';" + marginDeclarations + "}}</style><p>BodyOnly</p>";
}

using AngleSharp.Html.Parser;
using OfficeIMO;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Theory]
    [InlineData(false, "FFFFFF", "111111")]
    [InlineData(true, "14202C", "FFFFFF")]
    public void SemanticSlidesKeepSourceColorsReadableWhenTheReviewShellUsesTheOppositeTheme(bool dark, string background, string foreground) {
        using var bytes = new MemoryStream();
        using (var source = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = dark ? OfficeTheme.Dark : OfficeTheme.Modern })) {
            var slide = source.AddSlide();
            slide.BackgroundColor = background;
            slide.AddTextBox("Keep native colors").Paragraphs[0].Runs[0].Color = foreground;
            source.Save(bytes);
        }
        bytes.Position = 0;
        using var saved = PowerPointDocument.Load(bytes);
        var settings = new PowerPointHtmlSaveOptions();
        settings.DocumentOutput.DocumentTheme = dark ? OfficeTheme.Modern : OfficeTheme.Dark;
        string html = saved.ToHtml(settings);
        var parsed = new HtmlParser().ParseDocument(html);
        var slidePanel = Assert.Single(parsed.QuerySelectorAll("section.officeimo-slide"));
        string style = slidePanel.GetAttribute("style")!;
        Assert.Contains("--officeimo-page-background:#" + background + ";", style, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("background-color:var(--officeimo-page-background)", style, StringComparison.Ordinal);
        Assert.Contains("color:#" + foreground, slidePanel.InnerHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Keep native colors", slidePanel.TextContent, StringComparison.Ordinal);
        Assert.Contains("--officeimo-color-scheme:" + (dark ? "light" : "dark") + ";", parsed.QuerySelector("style")!.TextContent, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedDocumentThemeSurvivesCloneAndQuotesFontNamesInTheGeneratedShell(bool dark) {
        var basis = dark ? OfficeTheme.Dark : OfficeTheme.Modern;
        var theme = basis.WithPalette(basis.Palette.With(accent1: OfficeColor.FromRgb(0x12, 0x34, 0x56)))
            .WithFonts(basis.Fonts.With(body: new OfficeThemeFont("Review </style><script>bad()</script> \\\" Font", ["Aptos"])));
        var original = new OfficeHtmlDocumentOptions { DocumentTheme = theme, Title = "Theme review" };
        var copy = original.Clone();
        original.DocumentTheme = OfficeTheme.Modern.WithName("Another theme");
        string html = OfficeHtmlDocumentShell.WrapBody("<main><p>Visible review</p><table><tr><th>Owner</th></tr><tr><td>Operations</td></tr></table></main>", copy);
        var parsed = new HtmlParser().ParseDocument(html);
        Assert.Empty(parsed.QuerySelectorAll("script"));
        Assert.Single(parsed.QuerySelectorAll("style"));
        Assert.Equal("Visible review", parsed.QuerySelector("main p")!.TextContent);
        string css = parsed.QuerySelector("style")!.TextContent;
        Assert.Contains("--officeimo-accent:#123456;", css, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("--officeimo-page-background:" + basis.Palette.Background.ToCssColor() + ";", css, StringComparison.Ordinal);
        Assert.Contains("--officeimo-table-header-text:" + basis.TableStyle.HeaderText.Resolve(theme).ToCssColor() + ";", css, StringComparison.Ordinal);
        Assert.Contains("--officeimo-color-scheme:" + (dark ? "dark" : "light") + ";", css, StringComparison.Ordinal);
        Assert.Contains("\\3c ", css, StringComparison.Ordinal);
        Assert.Contains("\\22 ", css, StringComparison.Ordinal);
        Assert.Contains("\\5c ", css, StringComparison.Ordinal);
        copy.IncludeDefaultStyles = false;
        Assert.DoesNotContain("<style>", OfficeHtmlDocumentShell.WrapBody("<p>Unstyled</p>", copy), StringComparison.Ordinal);
    }
}

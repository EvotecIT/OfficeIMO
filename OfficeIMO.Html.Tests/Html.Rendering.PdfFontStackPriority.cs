using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public void HtmlPdf_InstalledCssStackKeepsItsFirstAvailableFamilyBeforeTheRendererDefault(
        bool customProperty, bool allowDocumentFonts) {
        string? primary = FindNonDefaultInstalledFamily();
        if (primary == null) return;
        string stack = "'" + primary + "',Arial,sans-serif";
        string css = customProperty
            ? ":root{--body-font:" + stack + "}p{font-family:var(--body-font,Arial,sans-serif)}"
            : "p{font-family:" + stack + "}";
        var options = new HtmlToPdfOptions();
        options.ResourcePolicy.AllowDocumentFontEmbedding = allowDocumentFonts;

        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<style>" + css + "</style><p>First family stays <strong>visible</strong></p>");
        HtmlPdfRenderResult result = HtmlPdfRenderedConverter.Convert(source, options);

        Assert.Equal(allowDocumentFonts, result.Document.Options.HasNamedFontFamily(primary));
        if (allowDocumentFonts) {
            HtmlRenderText text = Assert.Single(result.RenderResult!.Document.Pages[0].Visuals
                .OfType<HtmlRenderText>(), run => run.Text.Contains("First family", StringComparison.Ordinal));
            Assert.StartsWith(primary, text.Font.FamilyName, StringComparison.OrdinalIgnoreCase);
        }
        using var output = new MemoryStream();
        source.SaveAsPdf(output, options);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output.ToArray());
        Assert.All(independent.GetPage(1).Letters, letter => {
            Assert.DoesNotContain("__officeimo_", letter.FontName, StringComparison.Ordinal);
            Assert.DoesNotContain("OfficeIMO-SystemFace-", letter.FontName, StringComparison.Ordinal);
        });
        Assert.Contains("First family stays visible", independent.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("display:none")]
    [InlineData("visibility:hidden")]
    public void HtmlPdf_UnpaintedNamedTextDoesNotConsumeInstalledFontResources(string hiddenStyle) {
        string? primary = FindNonDefaultInstalledFamily();
        if (primary == null) return;
        var options = new HtmlToPdfOptions();
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            "<p style=\"font-family:'" + primary + "';" + hiddenStyle + "\">Hidden</p><p>Visible</p>"), options);

        Assert.DoesNotContain(rendered.Fonts.Faces,
            face => string.Equals(primary, face.FamilyName, StringComparison.OrdinalIgnoreCase));
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(),
            text => text.Text == "Visible");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void HtmlPdf_SuppliedPrefixCoverageAvoidsUnusedInstalledFontDiscovery(bool splitCoverage, bool list) {
        if (!PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Arial", out _)) return;
        byte[] first = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
            splitCoverage ? new[] { (int)'A', (int)'1', (int)'.', (int)' ' }
                : new[] { (int)'A', (int)'B', (int)'1', (int)'.', (int)' ' });
        var options = new HtmlToPdfOptions { MaxResourceBytes = first.Length * 2L };
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;
        options.Fonts.Add("SuppliedA", first);
        if (splitCoverage) options.Fonts.Add("SuppliedB", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('B'));
        HtmlConversionDocument source = HtmlConversionDocument.Parse(list
            ? "<ol style='font-family:SuppliedA,SuppliedB,Arial'><li>AB</li></ol>"
            : "<p style='font-family:SuppliedA,SuppliedB,Arial'>AB</p>");

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(source, options);

        rendered.RequireNoLoss();
        Assert.DoesNotContain(rendered.Fonts.Faces,
            face => string.Equals(face.FamilyName, "Arial", StringComparison.OrdinalIgnoreCase));
        using var output = new MemoryStream();
        OfficeConversionReport report = source.SaveAsPdf(output, options);
        report.ThrowIfLoss();
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output.ToArray());
        Assert.Contains("AB", independent.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_InstalledFamilyBeforeACoveringSuppliedPrefixKeepsItsCssPriority() {
        string? primary = FindNonDefaultInstalledFamily();
        if (primary == null) return;
        var options = new HtmlToPdfOptions();
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;
        options.Fonts.Add("Supplied", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B'));
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            "<p style=\"font-family:'" + primary + "',Supplied,Arial\">AB</p>"), options);

        HtmlRenderText text = Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>());
        Assert.StartsWith(primary, text.Font.FamilyName, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPdf_GlobalFallbackCountsAsPrefixCoverageOnlyWhenItIsInTheCssPrefix(bool explicitPrefix) {
        string? primary = FindNonDefaultInstalledFamily();
        if (primary == null) return;
        var options = new HtmlToPdfOptions();
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;
        options.Fonts.Add("SuppliedA", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A'));
        options.Fonts.Add("FallbackB", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('B'));
        options.Fonts.AddFallbackFamily("FallbackB");
        string prefix = explicitPrefix ? "SuppliedA,FallbackB," : "SuppliedA,";
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            "<p style=\"font-family:" + prefix + "'" + primary + "',Arial\">B</p>"), options);

        HtmlRenderText text = Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>());
        Assert.StartsWith(explicitPrefix ? "FallbackB" : primary, text.Font.FamilyName, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlPdf_CoveringSuppliedAliasWithALiteralCommaDoesNotDiscoverAnUnusedHostFamily() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B');
        var options = new HtmlToPdfOptions { MaxResourceBytes = data.Length * 2L };
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;
        options.Fonts.Add("Backing", data);
        options.Fonts.AddAlias("Supplied, body", "Backing");
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            "<p style=\"font-family:'Supplied, body',Arial\">AB</p>"), options);

        rendered.RequireNoLoss();
        Assert.DoesNotContain(rendered.Fonts.Faces,
            face => string.Equals(face.FamilyName, "Arial", StringComparison.OrdinalIgnoreCase));
    }

    private static string? FindNonDefaultInstalledFamily() =>
        new[] { "Aptos", "Helvetica Neue", "Liberation Serif", "DejaVu Serif" }
            .FirstOrDefault(name => PdfCore.PdfEmbeddedFontFamily.TryFromSystem(name, out _));

}

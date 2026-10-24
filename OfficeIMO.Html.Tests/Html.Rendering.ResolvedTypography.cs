using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlManagedTypography_ResolvedGlyphAdvanceMismatchCannotClaimNoLoss(bool nestedOverride) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120, includeSpace: true);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>html,body,p{margin:0;padding:0}p{font:20px '" + ManagedTextShapingTestAssets.FamilyName
            + "';font-feature-settings:'kern' 1}</style><p><bdo dir='rtl'>"
            + (nestedOverride ? "<span>AV</span>" : "AV") + "</bdo></p>");
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, ResolvedTypographyOptions(data));
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();

        Assert.True(rendered.Diagnostics.Any(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported),
            string.Join(" | ", text.Select(item => "Layout=" + item.TextAdvanceWidth
                + "; Prepared=" + item.PreparedShaping?.Measure(item.Font.Size))));
        Assert.Equal("AV", rendered.Text.Trim());
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void HtmlManagedTypography_ResolvedLatinOverrideKeepsItsPhysicalOrder(
        bool cssOverride, bool fontDirected) {
        byte[] data = fontDirected
            ? File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Drawing.Tests/TestAssets/OpenSans-Bold.ttf"))
            : ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(
                " abc\u202E\u202C".Select(scalar => (int)scalar).Distinct().ToDictionary(scalar => scalar, _ => (ushort)500),
                new[] { 32, 0x202E, 0x202C });
        string content = cssOverride ? "<bdo dir='rtl'>abc</bdo>" : "\u202Eabc\u202C";
        HtmlConversionDocument document = ResolvedTypographyDocument(content, "ltr", fontDirected);
        HtmlToPdfOptions options = ResolvedTypographyOptions(data, fontDirected);

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, options);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(data)!;
        int[] expected = "cba".Select(scalar => {
            Assert.True(font.TryGetGlyphMetrics(scalar, out int glyph, out _));
            return glyph;
        }).ToArray();

        Assert.Equal(expected, ObservePaintedTrueTypeGlyphs(rendered, text).Select(glyph => glyph.GlyphId));
        Assert.Equal("abc", rendered.Text.Trim());
        if (fontDirected) {
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Approximation);
            Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
        } else {
            rendered.RequireNoLoss();
            Assert.Contains(text, item => item.PreparedShaping != null);
            PdfCore.PdfDocumentConversionResult pdf = document.ToPdfDocumentResult(options);
            pdf.RequireNoLoss();
            Assert.Equal("abc", PdfCore.PdfReadDocument.Open(pdf.ToBytes()).ExtractText().Trim());
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult replay = Assert.Single(PdfCore.PdfImageExportExtensions.ExportImages(
                pdf, OfficeImageExportFormat.Png, new PdfCore.PdfImageExportOptions { Scale = 96D / 72D }));
            Assert.Equal(replay.Bytes, png.Bytes);
        }
    }

    [Theory]
    [InlineData("abc", "rtl", "abc")]
    [InlineData("<bdo dir='ltr'>אבג</bdo>", "rtl", "אבג")]
    public void HtmlManagedTypography_ResolvedRunsDistinguishParagraphAndOverrideDirection(
        string content, string paragraphDirection, string expectedPaint) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(
            " abcאבג".Select(scalar => (int)scalar).Distinct().ToDictionary(scalar => scalar, _ => (ushort)500),
            new[] { 32 });
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(
            ResolvedTypographyDocument(content, paragraphDirection), ResolvedTypographyOptions(data));
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(data)!;
        int[] expected = expectedPaint.Select(scalar => {
            Assert.True(font.TryGetGlyphMetrics(scalar, out int glyph, out _));
            return glyph;
        }).ToArray();

        Assert.Equal(expected, ObservePaintedTrueTypeGlyphs(rendered, text).Select(glyph => glyph.GlyphId));
        Assert.Contains(text, item => item.PreparedShaping != null);
        Assert.Equal(expectedPaint, rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    private static HtmlConversionDocument ResolvedTypographyDocument(string content, string direction, bool bold = false) =>
        HtmlConversionDocument.Parse("<style>@page{size:240px 100px;margin:0}html,body,p{margin:0;padding:0}"
            + "p{font:" + (bold ? "700 " : string.Empty) + "20px/24px '" + ManagedTextShapingTestAssets.FamilyName + "';font-feature-settings:'kern' 0}</style>"
            + "<p dir='" + direction + "'>" + content + "</p>");

    private static HtmlToPdfOptions ResolvedTypographyOptions(byte[] data, bool bold = false) {
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss,
            FontFamily = new PdfCore.PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName, data)
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, data, bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular);
        return options;
    }
}

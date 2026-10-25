using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Drawing;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using Xunit;
using PdfPig = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class MarkdownPdfFidelityTests {
    [Theory]
    [InlineData("日本語のテキスト。 中文文本。")]
    [InlineData("العربية עברית")]
    public void DefaultExportUsesMultilingualFallbacksAndPreservesCoveredLogicalText(string text) {
        Assert.True((new MarkdownToPdfOptions().TextFallbacks & PdfTextFallbackFeatures.MultilingualFonts) != 0);
        // Installed candidates establish available scalar coverage. Complete shaping
        // is verified separately with a pinned font rather than inferred from coverage.
        var fallback = new PdfOptions().UseTextFallbacks().EmbeddedFontFallbacks;
        if (fallback == null || !fallback.PlanText(text).IsFullyCovered) return;
        var conversion = MarkdownReader.Parse(text).ToPdfDocumentResult();
        byte[] bytes = conversion.ToBytes();
        Assert.NotEmpty(bytes);
        Assert.Contains(text, PdfReadDocument.Open(bytes).ExtractText());
    }

    [Fact]
    public void ExplicitNotoArabicFontPreservesLogicalTextWithoutUnsupportedShapingWarnings() {
        const string text = "مرحبا بالعالم";
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "NotoSansArabic-Regular.ttf"));
        var options = new MarkdownToPdfOptions {
            PdfOptions = new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, font, "Noto Arabic"),
            TextFallbacks = PdfTextFallbackFeatures.None,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };

        PdfDocumentConversionResult conversion = MarkdownReader.Parse(text).ToPdfDocumentResult(options);
        byte[] bytes = conversion.ToBytes();

        Assert.Equal(text, PdfReadDocument.Open(bytes).ExtractText().Trim());
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code is
            "unsupported-bidirectional-text-layout" or "unsupported-complex-script-shaping");
        conversion.RequireNoLoss();
    }

    [Fact]
    public void RtfDefaultExportUsesSharedMultilingualFallbacks() {
        var options = new RtfToPdfOptions();
        Assert.True((options.Clone().TextFallbacks & PdfTextFallbackFeatures.MultilingualFonts) != 0);
        const string text = "日本語のテキスト。 中文文本。";
        var fallback = new PdfOptions().UseTextFallbacks().EmbeddedFontFallbacks;
        if (fallback == null || !fallback.PlanText(text).IsFullyCovered) return;
        var document = RtfDocument.Create();
        document.AddParagraph(text);
        Assert.Contains("中文文本", PdfReadDocument.Open(document.ToPdfBytes()).ExtractText());
    }

    [Fact]
    public void CodeBlocksKeepLeadingSpacesTabsAndRepeatedWhitespaceWithoutChangingProse() {
        var document = MarkdownReader.Parse("```text\nROOT\n    SPACE  GAP\n\tTAB\n\t\tDEEP\n```\n\nProse   normal.");
        using var pdf = PdfPig.Open(document.ToPdfBytes(new MarkdownToPdfOptions { TextFallbacks = PdfTextFallbackFeatures.None }));
        var letters = pdf.GetPage(1).Letters;
        var root = letters.First(l => l.Value == "R");
        var space = letters.First(l => l.Value == "S");
        var tab = letters.First(l => l.Value == "T" && l.StartBaseLine.Y < root.StartBaseLine.Y);
        var deep = letters.First(l => l.Value == "D");
        Assert.True(space.StartBaseLine.X > root.StartBaseLine.X + 15D);
        Assert.True(tab.StartBaseLine.X > root.StartBaseLine.X + 20D);
        Assert.True(deep.StartBaseLine.X > tab.StartBaseLine.X + 20D);
        Assert.Contains("Courier", space.FontName);
        var codeLine = letters.Where(l => l.StartBaseLine.Y == space.StartBaseLine.Y && !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        var e = codeLine[4];
        var g = codeLine[5];
        Assert.True(g.StartBaseLine.X - e.StartBaseLine.X > e.Width * 2.5D);
    }

    [Theory]
    [InlineData("Body `WiMi` end.")]
    [InlineData("| H |\n| --- |\n| `WiMi` |")]
    public void InlineCodeUsesTheCodeBlockMonospaceFamily(string source) {
        using var pdf = PdfPig.Open(MarkdownReader.Parse(source).ToPdfBytes(new MarkdownToPdfOptions { TextFallbacks = PdfTextFallbackFeatures.None }));
        var code = pdf.GetPage(1).Letters.Where(l => "WiMi".Contains(l.Value) && l.Value.Length == 1).ToArray();
        Assert.Equal(4, code.Length);
        Assert.All(code, letter => Assert.Contains("Courier", letter.FontName));
        Assert.Equal(code[0].Width, code[1].Width, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FootnoteReferenceIsSuperscriptWithDestinationAndDefinitionFollowsBody(bool table) {
        string lead = table ? "| Header |\n| --- |\n| Body[^1]. |" : "Body[^1].";
        var result = MarkdownReader.Parse(lead + "\n\n[^1]: Data source: dashboard.\n\nTail.").ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();
        using var pdf = PdfPig.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        var marker = letters.First(l => l.Value == "1");
        var body = letters.First(l => l.Value == "B");
        Assert.True(marker.StartBaseLine.Y > body.StartBaseLine.Y);
        Assert.True(marker.FontSize < body.FontSize);
        string text = PdfReadDocument.Open(bytes).ExtractText();
        Assert.True(text.IndexOf("Tail.", StringComparison.Ordinal) < text.IndexOf("Data source:", StringComparison.Ordinal));
        Assert.DoesNotContain(result.Warnings, w => w.Code == "UnresolvedInternalLink");
        Assert.Contains("fn:1", System.Text.Encoding.ASCII.GetString(bytes));
        Assert.Contains("/GoTo", System.Text.Encoding.ASCII.GetString(bytes));
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Svg)]
    public void MarkdownPageImagesUseSerializedPdfPaginationAndGeometry(OfficeImageExportFormat format) {
        var conversion = MarkdownReader.Parse("# First\n\n```text\n    Indented\n```\n\n" + string.Join("\n\n", Enumerable.Repeat("A paragraph with `code` and a note[^1].", 30)) + "\n\n[^1]: Data source: dashboard.")
            .ToPdfDocumentResult(new MarkdownToPdfOptions { PdfOptions = new PdfOptions { PageWidth = 300, PageHeight = 300, MarginLeft = 20, MarginRight = 20, MarginTop = 20, MarginBottom = 20 } });
        var options = new PdfImageExportOptions { TargetDpi = 72 };
        var direct = conversion.ToImages(options).As(format).Export();
        var reopened = PdfReadDocument.Open(conversion.ToBytes());
        var saved = reopened.ToImages(options).As(format).Export();
        Assert.True(direct.Count > 1);
        Assert.Equal(reopened.Pages.Count, direct.Count);
        for (int i = 0; i < direct.Count; i++) {
            Assert.Equal(saved[i].Width, direct[i].Width);
            Assert.Equal(saved[i].Height, direct[i].Height);
            Assert.Equal(saved[i].Bytes, direct[i].Bytes);
        }
    }
}

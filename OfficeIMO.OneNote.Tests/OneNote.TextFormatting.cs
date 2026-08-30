using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OneNote;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OneNoteTextFormattingTests {
    [Fact]
    public void TransformTextCasePreservesNativeFormattingAndMetadata() {
        var run = new OneNoteTextRun {
            Text = "Hello WORLD",
            Hyperlink = "https://example.test",
            HyperlinkProtected = true
        };
        run.Style.FontFamily = "Aptos";
        run.Style.FontSize = 12.5;
        run.Style.ColorArgb = 0xFF336699;
        run.Style.HighlightColorArgb = 0xFFFFFF00;
        run.Style.Bold = true;
        run.Style.Italic = true;
        run.Style.Underline = true;
        run.Style.Strikethrough = true;
        run.Style.Superscript = true;
        run.Style.LanguageId = 1033;
        var opaque = new OneNoteOpaqueProperty { PropertyId = 0x12345678 };
        opaque.SetRawData(new byte[] { 1, 2, 3 });
        run.UnknownProperties.Add(opaque);

        OneNoteTextRun result = run.TransformTextCase(OfficeTextCase.ToggleCase, CultureInfo.InvariantCulture);

        Assert.Same(run, result);
        Assert.Equal("hELLO world", run.Text);
        Assert.Equal("Aptos", run.Style.FontFamily);
        Assert.Equal(12.5, run.Style.FontSize);
        Assert.Equal(0xFF336699U, run.Style.ColorArgb);
        Assert.Equal(0xFFFFFF00U, run.Style.HighlightColorArgb);
        Assert.True(run.Style.Bold);
        Assert.True(run.Style.Italic);
        Assert.True(run.Style.Underline);
        Assert.True(run.Style.Strikethrough);
        Assert.True(run.Style.Superscript);
        Assert.Equal(1033U, run.Style.LanguageId);
        Assert.Equal("https://example.test", run.Hyperlink);
        Assert.True(run.HyperlinkProtected);
        Assert.Single(run.UnknownProperties);
    }

    [Theory]
    [InlineData(true, false, OfficeTextBaseline.Superscript)]
    [InlineData(false, true, OfficeTextBaseline.Subscript)]
    [InlineData(false, false, OfficeTextBaseline.Normal)]
    public void RendererMapsNativeScriptFlagsToSharedRichTextBaseline(bool superscript, bool subscript, OfficeTextBaseline expected) {
        var run = new OneNoteTextRun { Text = "x" };
        run.Style.Superscript = superscript;
        run.Style.Subscript = subscript;
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(run);
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.2, Y = 0.5, Width = 2 } };
        outline.Children.Add(paragraph);
        var page = new OneNotePage { Title = "Script", PageSize = OneNotePageSize.IndexCard };
        page.Outlines.Add(outline);

        OneNotePageVisualSnapshot snapshot = OneNotePageRenderer.CreateSnapshot(page, new OneNotePageRenderingOptions { IncludeTitle = false });
        OfficeDrawingRichText richText = Assert.IsType<OfficeDrawingRichText>(snapshot.Drawing.Elements.Single(element => element is OfficeDrawingRichText));

        Assert.Equal(expected, Assert.Single(richText.Runs).Baseline);
    }

    [Fact]
    public void RendererPreservesFullyTransparentRunColors() {
        var run = new OneNoteTextRun { Text = "Invisible" };
        run.Style.ColorArgb = 0x00336699U;
        run.Style.HighlightColorArgb = 0x00FFF2CCU;
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(run);
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.2, Y = 0.5, Width = 2 } };
        outline.Children.Add(paragraph);
        var page = new OneNotePage { Title = "Alpha", PageSize = OneNotePageSize.IndexCard };
        page.Outlines.Add(outline);

        OfficeDrawingRichText richText = Assert.Single(page.ToDrawing(
            new OneNotePageRenderingOptions { IncludeTitle = false }).Elements.OfType<OfficeDrawingRichText>());
        OfficeRichTextRun actual = Assert.Single(richText.Runs);
        Assert.Equal((byte)0, actual.Color.A);
        Assert.Equal((byte)0, actual.BackgroundColor!.Value.A);
    }

    [Theory]
    [InlineData(true, false, OfficeTextBaseline.Superscript)]
    [InlineData(false, true, OfficeTextBaseline.Subscript)]
    public void RendererPreservesScriptBaselineForOrdinaryTextBesideInlineMath(bool superscript, bool subscript, OfficeTextBaseline expected) {
        var paragraph = new OneNoteParagraph();
        var scripted = new OneNoteTextRun { Text = "script" };
        scripted.Style.FontSize = 20D;
        scripted.Style.Superscript = superscript;
        scripted.Style.Subscript = subscript;
        paragraph.Runs.Add(scripted);
        paragraph.AddMath(OfficeMath.Identifier("x"));
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.2, Y = 0.5, Width = 3 } };
        outline.Children.Add(paragraph);
        var page = new OneNotePage { Title = "Mixed math", PageSize = OneNotePageSize.IndexCard };
        page.Outlines.Add(outline);

        OfficeDrawing drawing = page.ToDrawing(new OneNotePageRenderingOptions { IncludeTitle = false });
        OfficeDrawingRichText richText = Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>());
        OfficeRichTextRun run = Assert.Single(richText.Runs);

        Assert.Equal("script", run.Text);
        Assert.Equal(expected, run.Baseline);
        Assert.Equal(13D, run.EffectiveFontSize, 6);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void NativeTypographyExportsThroughEverySharedImageFormat(OfficeImageExportFormat format) {
        var run = new OneNoteTextRun { Text = "Styled" };
        run.Style.FontFamily = "Aptos";
        run.Style.FontSize = 16D;
        run.Style.ColorArgb = 0xFF336699U;
        run.Style.Bold = true;
        run.Style.Italic = true;
        run.Style.Underline = true;
        run.Style.Strikethrough = true;
        run.Style.Superscript = true;
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(run);
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.2, Y = 0.5, Width = 3 } };
        outline.Children.Add(paragraph);
        var page = new OneNotePage { Title = "Typography", PageSize = OneNotePageSize.IndexCard };
        page.Outlines.Add(outline);
        var section = new OneNoteSection { Name = "Typography" };
        section.Pages.Add(page);

        OfficeImageExportResult result = Assert.Single(section.ExportImages(format));

        Assert.Equal(format, result.Format);
        Assert.True(result.Bytes.Length > 32);
        if (format == OfficeImageExportFormat.Svg) {
            string svg = System.Text.Encoding.UTF8.GetString(result.Bytes);
            Assert.Contains("Styled", svg, System.StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("text-decoration=\"underline line-through\"", svg, System.StringComparison.OrdinalIgnoreCase);
        }
    }
}

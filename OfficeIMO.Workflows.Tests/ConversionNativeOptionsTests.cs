using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.PowerPoint;
using OfficeIMO.Rtf;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Workflows.Tests;

public sealed class ConversionNativeOptionsTests {
    [Fact]
    public async Task CanonicalMarkdownHtmlUsesReaderLimitsBeforeWriting() {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Long source marker"));
        using var output = new MemoryStream();
        ArgumentOutOfRangeException error = await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => OfficeConverter.ConvertAsync(input, OfficeDocumentFormat.Markdown,
            output, OfficeDocumentFormat.Html, new OfficeConversionOptions { MarkdownHtml = new MarkdownHtmlExportOptions {
                ReaderOptions = new MarkdownLoadOptions { MaxInputCharacters = 4 }
            } }));
        Assert.Equal("markdown", error.ParamName);
        Assert.Equal(0, output.Length);
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task HtmlRtfImportRetainsWordCssFontAndBaseDirectoryBehavior() {
        string directory = Directory.CreateTempSubdirectory("officeimo-html-rtf-import-").FullName;
        try {
            await File.WriteAllTextAsync(Path.Combine(directory, "style.css"), "p { color: #123456; }");
            var import = new HtmlWordExportOptions { BasePath = directory, FontFamily = "Courier New" };
            import.StylesheetPaths.Add("style.css");
            import.StylesheetContents.Add("p { font-size: 17pt; }");
            using var input = new MemoryStream(Encoding.UTF8.GetBytes("<p>Import marker</p>"));
            using var output = new MemoryStream();
            OfficeConversionReport report = await OfficeConverter.ConvertAsync(input, OfficeDocumentFormat.Html, output,
                OfficeDocumentFormat.Rtf, new OfficeConversionOptions { RegisteredSettings = new OfficeWorkflowRouteSettings { HtmlRtfImport = import } });
            output.Position = 0;
            using RtfDocument document = RtfDocument.Load(output);
            Assert.Contains(document.Fonts, font => font.Name == "Courier New");
            Assert.Contains(document.Colors, color => color.Red == 0x12 && color.Green == 0x34 && color.Blue == 0x56);
            Assert.Contains(document.Paragraphs.SelectMany(paragraph => paragraph.Runs), run => run.FontSize == 17);
            Assert.Contains(document.Paragraphs, paragraph => paragraph.ToPlainText().Contains("Import marker", StringComparison.Ordinal));
            Assert.Equal("Html", report.Source, ignoreCase: true);
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData(MarkdownSlideBreak.ThematicBreakOrHeading, 3)]
    [InlineData(MarkdownSlideBreak.ThematicBreak, 2)]
    [InlineData(MarkdownSlideBreak.Heading, 3)]
    public async Task MarkdownDeckOptionsControlRealSlideBoundariesFrontMatterAndNotes(MarkdownSlideBreak split, int slides) {
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("---\ntitle: Cover marker\nsubtitle: Subtitle marker\nauthor: Author marker\n---\n\n# First\n\nBody marker\n\n> Private speaker marker\n\n---\n\n# Second\n\n- List marker\n"));
        using var output = new MemoryStream();
        await OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown, output, OfficeDocumentFormat.PowerPoint,
            new OfficeConversionOptions { RegisteredSettings = new OfficeWorkflowRouteSettings {
                MarkdownPowerPoint = new MarkdownToPowerPointOptions { SlideBreak = split,
                    SlideSize = PowerPointSlideSizePreset.Screen4x3, RenderMermaidDiagrams = false }
            } });
        output.Position = 0;
        using PowerPointDocument presentation = PowerPointDocument.Load(output);
        Assert.Equal(slides, presentation.Slides.Count);
        Assert.Equal(10, presentation.SlideSize.Width.ToInches(), 5);
        string visible = string.Join("\n", presentation.Slides.SelectMany(slide => slide.TextBoxes).Select(box => box.Text));
        Assert.Contains("Cover marker", visible, StringComparison.Ordinal);
        Assert.Contains("Author marker", visible, StringComparison.Ordinal);
        Assert.DoesNotContain("Private speaker marker", visible, StringComparison.Ordinal);
        Assert.Contains(presentation.Slides, slide => slide.Notes.Text.Contains("Private speaker marker", StringComparison.Ordinal));
    }

    [Fact]
    public async Task MarkdownDeckCanKeepQuotesVisibleAndOmitFrontMatterWithAnotherTitleLevel() {
        using MarkdownDocument source = MarkdownDocument.Parse("---\ntitle: Omitted cover\n---\n\n# Body heading\n\n> Visible quote marker\n\n## Slide title marker\n\nContent marker");
        var converted = await source.ToPowerPointAsync(new MarkdownToPowerPointOptions { RenderFrontMatter = false,
            BlockquotesAsNotes = false, TitleHeadingLevel = 2, SlideBreak = MarkdownSlideBreak.Heading,
            RenderMermaidDiagrams = false });
        using PowerPointDocument presentation = converted.Value;
        Assert.Equal(2, presentation.Slides.Count);
        string visible = string.Join("\n", presentation.Slides.SelectMany(slide => slide.TextBoxes).Select(box => box.Text));
        Assert.DoesNotContain("Omitted cover", visible, StringComparison.Ordinal);
        Assert.Contains("Visible quote marker", visible, StringComparison.Ordinal);
        Assert.All(presentation.Slides, slide => Assert.True(string.IsNullOrEmpty(slide.Notes.Text)));
    }

    [Fact]
    public async Task MarkdownDeckRejectsDisallowedLocalImagesBeforeWriting() {
        string directory = Directory.CreateTempSubdirectory("officeimo-markdown-image-policy-").FullName;
        try {
            string path = Path.Combine(directory, "local.png");
            File.Copy(Path.Combine(AppContext.BaseDirectory, "samples", "provenance-demo.png"), path);
            using MarkdownDocument source = MarkdownDocument.Parse($"![Local marker]({path.Replace('\\', '/')})");
            using var output = new MemoryStream();
            OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(() => source.SaveAsPowerPointAsync(output,
                new MarkdownToPowerPointOptions { FailOnLoss = true }));
            Assert.Contains(error.Report.Diagnostics, finding => finding.Code == OfficeDiagnosticCodes.ImageSkippedByPolicy);
            Assert.Equal(0, output.Length);
            Assert.True(File.Exists(path));
        } finally { Directory.Delete(directory, recursive: true); }
    }

}

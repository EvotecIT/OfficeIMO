using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterThemesTests {
    private static OfficeTheme Theme() => OfficeTheme.Modern
        .WithTypeScale(OfficeTheme.Modern.TypeScale.With(body: OfficeLength.Points(21), heading1: OfficeLength.Points(29)))
        .WithTextStyle(OfficeTheme.Modern.TextStyle.With(headingColor: OfficeColor.Parse("#892345")));

    private static async Task<byte[]> Convert(byte[] source, OfficeDocumentFormat from, OfficeDocumentFormat to,
        OfficeConversionOptions settings) {
        using var input = new MemoryStream(source);
        using var output = new MemoryStream();
        await OfficeConverter.ConvertAsync(input, from, output, to, settings);
        Assert.True(input.CanRead);
        Assert.True(output.CanWrite);
        return output.ToArray();
    }

    private static double PdfSize(byte[] bytes, string text) {
        using var pdf = PdfDocument.Load(new MemoryStream(bytes));
        return Assert.Single(pdf.Read().TextBlocks, block => block.Text.Contains(text, StringComparison.Ordinal)).FontSize;
    }

    [Fact]
    public async Task CommonThemeReachesWordNativeXmlPdfAndHtmlWithoutChangingSource() {
        using var source = WordDocument.Create();
        source.AddParagraph("Native marker").SetFontSize(OfficeLength.Points(9));
        source.PageSetup.Margins.Left = OfficeLength.Points(39);
        byte[] bytes = source.ToBytes();
        OfficeTheme theme = Theme();
        var settings = new OfficeConversionOptions { Theme = theme,
            Word = new WordPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() } };
        byte[] pdf = await Convert(bytes, OfficeDocumentFormat.Word, OfficeDocumentFormat.Pdf, settings);
        Assert.Equal(21D, PdfSize(pdf, "Native marker"), 2);
        byte[] faithful = await Convert(bytes, OfficeDocumentFormat.Word, OfficeDocumentFormat.Pdf,
            new OfficeConversionOptions { Word = settings.Word });
        Assert.Equal(9D, PdfSize(faithful, "Native marker"), 2);
        string html = Encoding.UTF8.GetString(await Convert(bytes, OfficeDocumentFormat.Word, OfficeDocumentFormat.Html, settings));
        Assert.Contains("font-size:21", html, StringComparison.Ordinal);

        byte[] native = await Convert(bytes, OfficeDocumentFormat.Word, OfficeDocumentFormat.Word,
            new OfficeConversionOptions { Theme = theme, SourceName = "source.docx", TargetName = "themed.docm" });
        using var nativeWord = WordDocument.Load(new MemoryStream(native));
        Assert.Equal(OfficeLength.Points(21), nativeWord.Paragraphs.Single(paragraph => paragraph.Text == "Native marker").FontSize);
        Assert.Equal(OfficeLength.Points(39), nativeWord.PageSetup.Margins.Left);
        byte[] xml = await Convert(source.ExportFlatOpc(), OfficeDocumentFormat.Xml, OfficeDocumentFormat.Word,
            new OfficeConversionOptions { Theme = theme });
        using var xmlWord = WordDocument.Load(new MemoryStream(xml));
        Assert.Equal(OfficeLength.Points(21), xmlWord.Paragraphs.Single(paragraph => paragraph.Text == "Native marker").FontSize);
        Assert.Equal(OfficeLength.Points(9), source.Paragraphs.Single().FontSize);
    }

    [Fact]
    public async Task CommonThemeFillsMarkupOptionsAndKeepsExplicitAdapterAppearance() {
        byte[] markdown = Encoding.UTF8.GetBytes("# Heading\n\nMarkup marker");
        OfficeTheme theme = Theme();
        var settings = new OfficeConversionOptions { Theme = theme };
        string html = Encoding.UTF8.GetString(await Convert(markdown, OfficeDocumentFormat.Markdown, OfficeDocumentFormat.Html, settings));
        Assert.Contains("<!DOCTYPE html>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("--office-size-body: 21pt", html, StringComparison.Ordinal);
        Assert.Contains("var(--office-heading-color)", html, StringComparison.Ordinal);
        byte[] word = await Convert(markdown, OfficeDocumentFormat.Markdown, OfficeDocumentFormat.Word, settings);
        using var model = WordDocument.Load(new MemoryStream(word));
        Assert.Contains(model.Paragraphs, paragraph => paragraph.Text == "Markup marker");
        var normal = model.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Style>().Single(style => style.StyleId?.Value == "Normal");
        Assert.Equal("42", normal.StyleRunProperties!.FontSize!.Val!.Value);
        settings.Markdown = new MarkdownPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), TextFallbacks = PdfTextFallbackFeatures.None };
        Assert.Equal(21D, PdfSize(await Convert(markdown, OfficeDocumentFormat.Markdown, OfficeDocumentFormat.Pdf, settings), "Markup marker"), 2);
        settings.Markdown.Theme = MarkdownVisualTheme.FromOfficeTheme(theme.WithTypeScale(theme.TypeScale.With(body: OfficeLength.Points(12))));
        Assert.Equal(12D, PdfSize(await Convert(markdown, OfficeDocumentFormat.Markdown, OfficeDocumentFormat.Pdf, settings), "Markup marker"), 2);
        Assert.Null(settings.MarkdownWord);
        Assert.Null(settings.MarkdownHtml);

        byte[] authoredHtml = Encoding.UTF8.GetBytes("<h1 style=\"font-size:17pt\">Authored marker</h1><p>Common marker</p>");
        settings.Html = new HtmlPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        byte[] htmlPdf = await Convert(authoredHtml, OfficeDocumentFormat.Html, OfficeDocumentFormat.Pdf, settings);
        Assert.Equal(17D, PdfSize(htmlPdf, "Authored marker"), 1);
        Assert.Equal(21D, PdfSize(htmlPdf, "Common marker"), 1);
        byte[] htmlWord = await Convert(authoredHtml, OfficeDocumentFormat.Html, OfficeDocumentFormat.Word, settings);
        using var fromHtml = WordDocument.Load(new MemoryStream(htmlWord));
        Assert.Equal(OfficeLength.Points(17), fromHtml.Paragraphs.Single(paragraph => paragraph.Text == "Authored marker").FontSize);

        using var template = WordDocument.Create();
        template.PageSetup.Margins.Left = OfficeLength.Points(19);
        settings.MarkdownWord = new MarkdownWordExportOptions { Template = template, FontFamily = "Verdana" };
        byte[] templatedBytes = await Convert(markdown, OfficeDocumentFormat.Markdown, OfficeDocumentFormat.Word, settings);
        using var templated = WordDocument.Load(new MemoryStream(templatedBytes));
        Assert.Equal(OfficeLength.Points(19), templated.PageSetup.Margins.Left);
        Assert.Equal("Verdana", templated.Paragraphs.Single(paragraph => paragraph.Text == "Markup marker").FontFamily);
        template.AddParagraph("Still caller-owned");
    }

    [Fact]
    public async Task UnsupportedThemesAndHtmlWithoutCssAreRejectedBeforeDestinationWrites() {
        using var document = WordDocument.Create();
        document.AddParagraph("Preserved source");
        using var word = new MemoryStream(document.ToBytes());
        using var output = new MemoryStream(Encoding.UTF8.GetBytes("Keep destination"));
        var settings = new OfficeConversionOptions { Theme = Theme() };
        NotSupportedException error = await Assert.ThrowsAsync<NotSupportedException>(() => OfficeConverter.ConvertAsync(word,
            OfficeDocumentFormat.Word, output, OfficeDocumentFormat.Markdown, settings));
        Assert.Contains("docx-markdown", error.Message, StringComparison.Ordinal);
        Assert.Equal("Keep destination", Encoding.UTF8.GetString(output.ToArray()));
        using var markdown = new MemoryStream(Encoding.UTF8.GetBytes("# Fragment"));
        settings.MarkdownHtml = new MarkdownHtmlExportOptions { Kind = MarkdownHtmlKind.Fragment };
        await Assert.ThrowsAsync<NotSupportedException>(() => OfficeConverter.ConvertAsync(markdown,
            OfficeDocumentFormat.Markdown, output, OfficeDocumentFormat.Html, settings));
        Assert.Equal("Keep destination", Encoding.UTF8.GetString(output.ToArray()));
        settings.MarkdownHtml.IncludeCssInFragment = true;
        byte[] fragment = await Convert(markdown.ToArray(), OfficeDocumentFormat.Markdown, OfficeDocumentFormat.Html, settings);
        Assert.Contains("<style", Encoding.UTF8.GetString(fragment), StringComparison.Ordinal);
        Assert.DoesNotContain("<!DOCTYPE", Encoding.UTF8.GetString(fragment), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ThemeSnapshotsAndRuntimeBatchesPreserveAppearanceWhilePersistenceRejectsIt() {
        string root = Directory.CreateTempSubdirectory("officeimo-theme-batch-").FullName;
        try {
            OfficeTheme theme = Theme();
            var options = new OfficeConversionOptions { Theme = theme };
            Assert.Same(theme, options.Clone().Theme);
            Assert.Same(theme, options.ForRoute("markdown-html").Theme);
            Assert.Throws<NotSupportedException>(() => options.ForRoute("docx-markdown"));
            string source = Path.Combine(root, "source.md");
            File.WriteAllText(source, "# Runtime theme\n\nBatch marker");
            var batch = new OfficeBatchOptions { InputPaths = [source], OutputDirectory = Path.Combine(root, "ordinary"),
                TargetExtension = ".html", ConversionOptions = options };
            Assert.Throws<NotSupportedException>(() => batch.ToJson(root));
            var checkpointed = batch with { OutputDirectory = Path.Combine(root, "held"), ResumeFrom = Path.Combine(root, "state", "plan.json") };
            OfficeBatchPlan plan = await OfficeConverter.PlanAsync(checkpointed);
            Assert.False(plan.CanRun);
            await Assert.ThrowsAsync<NotSupportedException>(() => OfficeConverter.ConvertAsync(checkpointed));
            Assert.False(Directory.Exists(checkpointed.OutputDirectory));
            Assert.False(Directory.Exists(Path.GetDirectoryName(checkpointed.ResumeFrom)));
            OfficeBatchReport report = await OfficeConverter.ConvertAsync(batch);
            Assert.Equal(1, report.Summary.Completed);
            string artifact = File.ReadAllText(Assert.Single(report.Items).OutputPath!);
            Assert.Contains("--office-size-body: 21pt", artifact, StringComparison.Ordinal);
            Assert.Same(theme, options.Theme);
        } finally { Directory.Delete(root, true); }
    }
}

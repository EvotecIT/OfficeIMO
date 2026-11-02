using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterTemplateLifetimeTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MarkdownTemplateRemainsCallerOwnedAfterPublicationOrOutputLimitRejection(bool rejectOutput) {
        using var template = WordDocument.Create();
        var retained = template.AddParagraph("Template prefix");
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("Imported template evidence"));
        using var output = new MemoryStream();
        var settings = new OfficeConversionOptions { MarkdownWord = new WordMarkdownImportOptions { Template = template } };
        if (rejectOutput) settings.MaximumOutputBytes = 1;
        if (rejectOutput) {
            await Assert.ThrowsAsync<OfficeConversionException>(() => OfficeConverter.ConvertAsync(source,
                OfficeDocumentFormat.Markdown, output, OfficeDocumentFormat.Word, settings));
        } else {
            await OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown, output, OfficeDocumentFormat.Word, settings);
            output.Position = 0;
            using var published = WordDocument.Load(output);
            Assert.Contains(published.Paragraphs, paragraph => paragraph.Text.Contains("Imported template evidence", StringComparison.Ordinal));
        }
        var added = template.AddParagraph("Caller still owns the template");
        Assert.Contains(retained, template.Paragraphs);
        Assert.Contains(added, template.Paragraphs);
        Assert.Equal("Caller still owns the template", added.Text);
        Assert.Same(template, settings.MarkdownWord.Template);
    }
}

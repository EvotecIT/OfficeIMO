using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Save_GeneratedTableOfContentsStoresEntriesAndDefinesItsStyles() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create()) {
            document.AddTableOfContent();
            document.AddParagraph("Introduction").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Body");
            document.AddParagraph("Details").Style = WordParagraphStyles.Heading2;
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument loaded = WordDocument.Load(stream);
        var body = loaded._wordprocessingDocument.MainDocumentPart!.Document.Body!;
        var entryStyles = body.Descendants<SdtBlock>().SelectMany(block => block.Descendants<ParagraphStyleId>())
            .Select(style => style.Val?.Value).ToList();
        Assert.Contains("TOC1", entryStyles);
        Assert.Contains("TOC2", entryStyles);
        string tocText = string.Concat(body.Descendants<SdtBlock>().First().Descendants<Text>().Select(text => text.Text));
        Assert.Contains("Introduction", tocText);
        Assert.Contains("Details", tocText);

        var styleIds = loaded._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Select(style => style.StyleId?.Value).ToHashSet();
        Assert.Contains("TOCHeading", styleIds);
        Assert.Contains("TOC1", styleIds);
        Assert.Contains("TOC3", styleIds);
    }
}

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeDocStyleDiscoveryRejectsDuplicateRunningStoryReferencesBeforeScanningParts(bool footer) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        MainDocumentPart main = document._wordprocessingDocument.MainDocumentPart!;
        Paragraph styled = new(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), new Run(new Text("Story")));
        OpenXmlPart part;
        if (footer) { FooterPart story = main.AddNewPart<FooterPart>(); story.Footer = new Footer(styled); part = story; }
        else { HeaderPart story = main.AddNewPart<HeaderPart>(); story.Header = new Header(styled); part = story; }
        Style style = main.StyleDefinitionsPart!.Styles!.Elements<Style>().Single(item => item.StyleId == "Heading1");
        style.StyleRunProperties ??= new StyleRunProperties();
        style.StyleRunProperties.AddChild(new Position { Val = "2" }, true);
        SectionProperties section = main.Document.Body!.Elements<SectionProperties>().Single();
        if (footer) {
            var reference = new FooterReference { Id = main.GetIdOfPart(part), Type = HeaderFooterValues.Default };
            section.Append(reference);
            for (int index = 0; index < 128; index++) section.Append((FooterReference)reference.CloneNode(true));
        } else {
            var reference = new HeaderReference { Id = main.GetIdOfPart(part), Type = HeaderFooterValues.Default };
            section.Append(reference);
            for (int index = 0; index < 128; index++) section.Append((HeaderReference)reference.CloneNode(true));
        }

        NotSupportedException error = Assert.Throws<NotSupportedException>(() => document.ToBytes(WordFileFormat.Doc));
        Assert.Contains("multiple", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeDocStyleDiscoveryIgnoresUnwrittenNoteSeparators(bool endnote) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        MainDocumentPart main = document._wordprocessingDocument.MainDocumentPart!;
        Paragraph empty = new(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }));
        if (endnote) (main.EndnotesPart ?? main.AddNewPart<EndnotesPart>()).Endnotes = new Endnotes(
            new Endnote(empty) { Id = 0, Type = FootnoteEndnoteValues.Separator });
        else (main.FootnotesPart ?? main.AddNewPart<FootnotesPart>()).Footnotes = new Footnotes(
            new Footnote(empty) { Id = 0, Type = FootnoteEndnoteValues.Separator });
        Style style = main.StyleDefinitionsPart!.Styles!.Elements<Style>().Single(item => item.StyleId == "Heading1");
        style.StyleRunProperties ??= new StyleRunProperties();
        style.StyleRunProperties.AddChild(new Position { Val = "2" }, true);
        Assert.Empty(document.ValidateDocument());
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc })));
        Assert.Equal("Body", loaded.Paragraphs.Single().Text);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void NativeDocStyleDiscoveryIgnoresEmptyUnreferencedHeaderFooterParts(bool footer, bool referenced) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        MainDocumentPart main = document._wordprocessingDocument.MainDocumentPart!;
        Paragraph empty = new(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }));
        OpenXmlPart part;
        if (footer) { FooterPart story = main.AddNewPart<FooterPart>(); story.Footer = new Footer(empty); part=story; }
        else { HeaderPart story = main.AddNewPart<HeaderPart>(); story.Header = new Header(empty); part=story; }
        if (referenced) {
            SectionProperties section = main.Document.Body!.Elements<SectionProperties>().Single();
            if (footer) section.AddChild(new FooterReference { Id = main.GetIdOfPart(part), Type = HeaderFooterValues.Default }, true);
            else section.AddChild(new HeaderReference { Id = main.GetIdOfPart(part), Type = HeaderFooterValues.Default }, true);
        }
        Style style = main.StyleDefinitionsPart!.Styles!.Elements<Style>().Single(item => item.StyleId == "Heading1");
        style.StyleRunProperties ??= new StyleRunProperties();
        style.StyleRunProperties.AddChild(new Position { Val = "2" }, true);
        Assert.Empty(document.ValidateDocument());
        if (referenced) {
            NotSupportedException error = Assert.Throws<NotSupportedException>(() => document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }));
            Assert.Contains("position", error.Message, StringComparison.Ordinal);
            return;
        }
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc })));
        Assert.Equal("Body", loaded.Paragraphs.Single().Text);
    }
}

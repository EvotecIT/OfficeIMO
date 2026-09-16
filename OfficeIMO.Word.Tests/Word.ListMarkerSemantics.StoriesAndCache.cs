using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class WordListMarkerSemanticsTests {
    [Fact]
    public void NumberingCountersAreIndependentAcrossDocumentStories() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));

        WordParagraph bodyFirst = document.AddParagraph("Body first");
        AttachToList(bodyFirst, list.NumberId);
        WordParagraph bodySecond = document.AddParagraph("Body second");
        AttachToList(bodySecond, list.NumberId);

        WordParagraph headerFirst = document.Header!.Default!.AddParagraph("Header first");
        AttachToList(headerFirst, list.NumberId);
        WordParagraph headerSecond = document.Header.Default.AddParagraph("Header second");
        AttachToList(headerSecond, list.NumberId);
        WordParagraph firstPageHeader = document.HeaderFirstOrCreate.AddParagraph("First-page header");
        AttachToList(firstPageHeader, list.NumberId);
        WordParagraph evenPageHeader = document.HeaderEvenOrCreate.AddParagraph("Even-page header");
        AttachToList(evenPageHeader, list.NumberId);
        WordParagraph footer = document.Footer!.Default!.AddParagraph("Footer item");
        AttachToList(footer, list.NumberId);

        var markers = WordDocumentTraversal.BuildListMarkers(document);
        Assert.Equal("1.", markers[bodyFirst].Marker);
        Assert.Equal("2.", markers[bodySecond].Marker);
        Assert.Equal("1.", markers[headerFirst].Marker);
        Assert.Equal("2.", markers[headerSecond].Marker);
        Assert.Equal("1.", markers[firstPageHeader].Marker);
        Assert.Equal("1.", markers[evenPageHeader].Marker);
        Assert.Equal("1.", markers[footer].Marker);

        var indices = WordDocumentTraversal.BuildListIndices(document);
        Assert.Equal(2, indices[bodySecond].Index);
        Assert.Equal(2, indices[headerSecond].Index);
        Assert.Equal(1, indices[firstPageHeader].Index);
        Assert.Equal(1, indices[evenPageHeader].Index);
        Assert.Equal(1, indices[footer].Index);
    }

    [Fact]
    public void PdfHeaderAndBodyListsEachStartAtOne() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        WordParagraph body = document.AddParagraph("Body item");
        AttachToList(body, list.NumberId);
        WordParagraph header = document.Header!.Default!.AddParagraph("Header item");
        AttachToList(header, list.NumberId);

        string pdfText = PdfReadDocument.Open(document.ToPdfBytes()).ExtractText();
        Assert.Contains("1. Header item", pdfText, StringComparison.Ordinal);
        Assert.Contains("1. Body item", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("2. Header item", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void RemovingLevelInvalidatesStyleLinkedListMembership() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        list.Numbering.Levels[0].OpenXmlElement.Append(new ParagraphStyleIdInLevel { Val = "Issue2510RemovableLevel" });
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "Issue2510RemovableLevel" });
        WordParagraph item = document.AddParagraph("Linked item");
        item._paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Issue2510RemovableLevel" });

        Assert.True(item.IsListItem);
        Assert.Equal("1.", WordDocumentTraversal.BuildListMarkers(document)[item].Marker);
        list.Numbering.Levels[0].Remove();
        Assert.False(item.IsListItem);
        Assert.DoesNotContain(item, WordDocumentTraversal.BuildListMarkers(document).Keys);
    }

    [Fact]
    public void ReplacingAbstractDefinitionInvalidatesStyleLinkedListMembership() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        list.Numbering.Levels[0].OpenXmlElement.Append(new ParagraphStyleIdInLevel { Val = "Issue2510ReplaceDefinition" });
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "Issue2510ReplaceDefinition" });
        WordParagraph item = document.AddParagraph("Linked item");
        item._paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Issue2510ReplaceDefinition" });

        Assert.True(item.IsListItem);
        list.ConvertToBulleted();
        Assert.False(item.IsListItem);
        Assert.DoesNotContain(item, WordDocumentTraversal.BuildListMarkers(document).Keys);
    }
}

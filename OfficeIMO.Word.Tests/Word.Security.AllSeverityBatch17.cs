using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = OfficeIMO.Drawing.OfficeColor;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LineColorFallsBackForMissingAutomaticAndMalformedValues() {
        using WordDocument document = WordDocument.Create();
        WordLine line = document.AddParagraph("line").AddLine(0, 0, 20, 0);

        line._line.StrokeColor = "auto";
        Assert.Equal(Color.Black, line.Color);
        line._line.StrokeColor = "not-a-color";
        Assert.Equal(Color.Black, line.Color);
        line._line.StrokeColor = null;
        Assert.Equal(Color.Black, line.Color);
    }

    [Fact]
    public void CloneListBeforeParagraphPreservesSourceOrderInDocumentXml() {
        using WordDocument document = WordDocument.Create();
        WordParagraph marker = document.AddParagraph("marker");
        WordList list = document.AddList(WordListStyle.Bulleted);
        list.AddItem("first");
        list.AddItem("second");

        _ = list.Clone(marker, after: false);

        string[] paragraphs = document.Paragraphs.Select(paragraph => paragraph.Text).ToArray();
        int markerIndex = Array.IndexOf(paragraphs, "marker");
        Assert.True(markerIndex >= 2);
        Assert.Equal("first", paragraphs[markerIndex - 2]);
        Assert.Equal("second", paragraphs[markerIndex - 1]);
    }

    [Fact]
    public void RemovingSectionDoesNotDeleteSameNumberingItemsOutsideTheSection() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListStyle.Bulleted);
        WordParagraph outside = list.AddItem("outside-list-item");
        WordSection targetSection = document.AddSection();
        WordParagraph target = targetSection.AddParagraph("target-list-item");
        target._paragraph.ParagraphProperties ??= new ParagraphProperties();
        target._paragraph.ParagraphProperties.NumberingProperties =
            (NumberingProperties)outside._paragraph.ParagraphProperties!.NumberingProperties!.CloneNode(true);

        Assert.NotEmpty(targetSection.Lists);
        targetSection.RemoveSection();

        Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "outside-list-item");
        Assert.DoesNotContain(document.Paragraphs, paragraph => paragraph.Text == "target-list-item");
    }

    [Fact]
    public void RemovingSectionCleansNumberingUsedOnlyInTablesAndHeaders() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("retained");
        WordSection targetSection = document.AddSection();

        WordList tableList = targetSection.AddTable(1, 1).Rows[0].Cells[0].AddList(WordListStyle.Bulleted);
        tableList.AddItem("table-list-item");
        targetSection.AddHeadersAndFooters();
        WordHeaderFooter header = Assert.IsAssignableFrom<WordHeaderFooter>(targetSection.Header.Default);
        WordList headerList = header.AddList(WordListStyle.Numbered);
        headerList.AddItem("header-list-item");
        int[] removedNumberingIds = { tableList._numberId, headerList._numberId };

        targetSection.RemoveSection();

        Numbering? numbering = document._wordprocessingDocument.MainDocumentPart?
            .NumberingDefinitionsPart?
            .Numbering;
        Assert.DoesNotContain(
            numbering?.Elements<NumberingInstance>() ?? Enumerable.Empty<NumberingInstance>(),
            instance => removedNumberingIds.Contains(instance.NumberID?.Value ?? -1));
    }

    [Fact]
    public void RemovingSectionPreservesNumberingReferencedByFootnotes() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("retained");
        WordSection targetSection = document.AddSection();
        WordList removedList = targetSection.AddParagraph("section-list").AddList(WordListStyle.Bulleted);
        int sharedNumberingId = removedList._numberId;

        MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
        FootnotesPart footnotesPart = mainPart.FootnotesPart ?? mainPart.AddNewPart<FootnotesPart>();
        footnotesPart.Footnotes = new Footnotes(
            new Footnote(
                new Paragraph(
                    new ParagraphProperties(
                        new NumberingProperties(new NumberingId { Val = sharedNumberingId })),
                    new Run(new Text("footnote-list-item")))) {
                Id = 1
            });

        targetSection.RemoveSection();

        Numbering numbering = Assert.IsType<Numbering>(
            mainPart.NumberingDefinitionsPart?.Numbering);
        Assert.Contains(
            numbering.Elements<NumberingInstance>(),
            instance => instance.NumberID?.Value == sharedNumberingId);
    }

    [Fact]
    public void RemoveSectionToleratesMissingAndDuplicateHeaderRelationships() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("first section");
        WordSection section = document.AddSection();
        section.AddHeadersAndFooters();
        HeaderReference valid = section._sectionProperties.Elements<HeaderReference>().First();
        section._sectionProperties.Append((HeaderReference)valid.CloneNode(true));
        section._sectionProperties.Append(new HeaderReference {
            Type = HeaderFooterValues.Even,
            Id = "rIdMissingHeader"
        });

        Exception? exception = Record.Exception(section.RemoveSection);

        Assert.Null(exception);
        Assert.Single(document.Sections);
    }

    [Fact]
    public void SelectiveHeaderFooterRemovalToleratesMissingAndDuplicateRelationships() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        SectionProperties section = document.Sections[0]._sectionProperties;
        HeaderReference header = section.Elements<HeaderReference>().First(reference => reference.Type?.Value == HeaderFooterValues.Default);
        FooterReference footer = section.Elements<FooterReference>().First(reference => reference.Type?.Value == HeaderFooterValues.Default);
        section.Append((HeaderReference)header.CloneNode(true));
        section.Append((FooterReference)footer.CloneNode(true));
        section.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = "rIdMissingHeader" });
        section.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = "rIdMissingFooter" });

        Exception? exception = Record.Exception(() => {
            WordHeader.RemoveHeaders(document._wordprocessingDocument, WordHeaderFooterType.Default);
            WordFooter.RemoveFooters(document._wordprocessingDocument, WordHeaderFooterType.Default);
        });

        Assert.Null(exception);
        Assert.Empty(section.Elements<HeaderReference>().Where(reference => reference.Type?.Value == HeaderFooterValues.Default));
        Assert.Empty(section.Elements<FooterReference>().Where(reference => reference.Type?.Value == HeaderFooterValues.Default));
    }

    [Fact]
    public void RemovingShapePreservesTextInItsParagraph() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("keep this text");
        WordShape shape = paragraph.AddShape(20, 10);

        shape.Remove();

        Assert.Equal("keep this text", paragraph.Text);
        Assert.False(paragraph.IsShape);
    }

    [Fact]
    public void SetFixedWidthAcceptsRowsWhoseCellsWereRemoved() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].Remove();

        Exception? exception = Record.Exception(() => table.SetFixedWidth(50));

        Assert.Null(exception);
    }

    [Fact]
    public void ListFormattingReturnsNullForAutomaticAndMalformedColors() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListStyle.Bulleted);

        list.ColorHex = "auto";
        Assert.Null(list.Color);
        list.ColorHex = "not-a-color";
        Assert.Null(list.Color);
    }

    [Fact]
    public void DotxConversionRemovesAttachedTemplateRelationships() {
        string templatePath = Path.Combine(_directoryDocuments, "ExampleTemplate.dotx");
        string outputPath = Path.Combine(_directoryWithFiles, "ExampleTemplate_NoAttachedTemplate.docx");

        WordHelpers.ConvertDotxToDocx(templatePath, outputPath);

        using WordprocessingDocument converted = WordprocessingDocument.Open(outputPath, false);
        DocumentSettingsPart? settingsPart = converted.MainDocumentPart?.DocumentSettingsPart;
        Assert.DoesNotContain(settingsPart?.ExternalRelationships ?? Enumerable.Empty<ExternalRelationship>(),
            relationship => relationship.RelationshipType.EndsWith("/attachedTemplate", StringComparison.Ordinal));
        Assert.Empty(settingsPart?.Settings?.Elements<AttachedTemplate>() ?? Enumerable.Empty<AttachedTemplate>());
    }
}

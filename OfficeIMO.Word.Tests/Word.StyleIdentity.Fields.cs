using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void RefreshingTocUsesCaseDistinctOutlineBaseStyle() {
        Style child = new(new StyleName { Val = "Outline child" }, new BasedOn { Val = "caseOutline" }) {
            Type = StyleValues.Paragraph,
            StyleId = "CaseOutline"
        };
        Style parent = new(new StyleName { Val = "Outline base" },
            new StyleParagraphProperties(new OutlineLevel { Val = 1 })) {
            Type = StyleValues.Paragraph,
            StyleId = "caseOutline"
        };
        Style body = new(new StyleName { Val = "Authored body style" }) {
            Type = StyleValues.Paragraph,
            StyleId = "heading1"
        };
        Style authoredToc = new(new StyleName { Val = "Authored lowercase TOC" },
            new StyleRunProperties(new Color { Val = "315F23" })) {
            Type = StyleValues.Paragraph,
            StyleId = "toc2"
        };
        string authoredTocXml = authoredToc.OuterXml;
        using var input = CreateCaseDistinctStyleDocument(child, parent, body, authoredToc);
        using WordDocument document = WordDocument.Load(input);
        document.AddParagraph("Authored lowercase body").SetStyleId("heading1");
        WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 3);

        WordTableOfContentRefreshReport report = toc.RefreshEntries();

        Assert.Equal("Minimal document", Assert.Single(report.Entries).Text);
        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        MainDocumentPart mainPart = saved.MainDocumentPart!;
        Paragraph entry = Assert.Single(mainPart.Document.Descendants<Paragraph>(), paragraph =>
            paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "TOC2");
        Assert.Contains("Minimal document", entry.InnerText);
        Style[] styles = mainPart.StyleDefinitionsPart!.Styles!.Elements<Style>().ToArray();
        Assert.Equal(authoredTocXml, Assert.Single(styles, style => style.StyleId == "toc2").OuterXml);
        Style generatedToc = Assert.Single(styles, style => style.StyleId == "TOC2");
        TabStop leader = Assert.Single(generatedToc.StyleParagraphProperties!.Tabs!.Elements<TabStop>());
        Assert.Equal(TabStopValues.Right, leader.Val!.Value);
        Assert.Equal(TabStopLeaderCharValues.Dot, leader.Leader!.Value);
        Assert.Equal("caseOutline", mainPart.StyleDefinitionsPart!.Styles!.Elements<Style>()
            .Single(style => style.StyleId == "CaseOutline").BasedOn!.Val!.Value);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2013).Validate(saved));
    }

    [Fact]
    public void UpdatingRefNumberUsesCaseDistinctNumberedBaseStyle() {
        Style child = new(new StyleName { Val = "Number child" }, new BasedOn { Val = "caseNumber" }) {
            Type = StyleValues.Paragraph,
            StyleId = "CaseNumber"
        };
        Style parent = new(new StyleName { Val = "Number base" }, new StyleParagraphProperties(
            new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 }))) {
            Type = StyleValues.Paragraph,
            StyleId = "caseNumber"
        };
        using var input = CreateCaseDistinctStyleDocument(child, parent);
        using (WordprocessingDocument package = WordprocessingDocument.Open(input, true)) {
            NumberingDefinitionsPart numbering = package.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();
            numbering.Numbering = new Numbering(
                new AbstractNum(new Level(new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = 0 },
                new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
        input.Position = 0;
        using WordDocument document = WordDocument.Load(input);
        Assert.Single(document.Paragraphs).AddBookmark("CaseNumberTarget");
        document.AddParagraph("Reference: ")._paragraph.Append(new SimpleField(new Run(new Text("stale"))) {
            Instruction = " REF CaseNumberTarget \\n "
        });

        WordFieldUpdateResult result = Assert.Single(document.UpdateFieldsAndGetReport().Results);

        Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
        Assert.Equal("1", result.ResultText);
        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        Assert.Equal("1", Assert.Single(saved.MainDocumentPart!.Document.Descendants<SimpleField>()).InnerText);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    [Fact]
    public void UpdatingSequenceResetIgnoresCaseDistinctBodyStyle() {
        Style chapter = new(new StyleName { Val = "Authored chapter" }) {
            Type = StyleValues.Paragraph,
            StyleId = "Heading1"
        };
        Style body = new(new StyleName { Val = "Authored body" }) {
            Type = StyleValues.Paragraph,
            StyleId = "heading1"
        };
        using var input = CreateCaseDistinctStyleDocument(chapter, body);
        using WordDocument document = WordDocument.Load(input);
        document.AddParagraph("Figure ").AddField(WordFieldType.Seq,
            parameters: new List<string> { "Figure", "\\s 1" });
        document.AddParagraph("Authored lowercase body").SetStyleId("heading1");
        document.AddParagraph("Figure ").AddField(WordFieldType.Seq,
            parameters: new List<string> { "Figure", "\\s 1" });
        document.AddParagraph("Chapter two").SetStyleId("Heading1");
        document.AddParagraph("Figure ").AddField(WordFieldType.Seq,
            parameters: new List<string> { "Figure", "\\s 1" });

        WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

        Assert.Equal(3, report.UpdatedCount);
        Assert.Equal(new[] { "1", "2", "1" }, report.Results.Select(result => result.ResultText));
        byte[] savedBytes = document.ToBytes();
        using var savedStream = new MemoryStream(savedBytes, writable: false);
        using (WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false)) {
            Assert.Empty(new OpenXmlValidator().Validate(saved));
        }
        using var reloadedStream = new MemoryStream(savedBytes, writable: false);
        using WordDocument reloaded = WordDocument.Load(reloadedStream);
        Assert.Equal(new[] { "1", "2", "1" }, reloaded.InspectFields()
            .Where(field => field.FieldType == WordFieldType.Seq).Select(field => field.ResultText));
    }
}

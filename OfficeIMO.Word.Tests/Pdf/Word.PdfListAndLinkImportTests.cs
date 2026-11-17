using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf.Word;
using Xunit;
using PdfCore = OfficeIMO.Pdf;
using OfficeWordDocument = OfficeIMO.Word.WordDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void PdfSemanticImport_KeepsBulletAndNumberedListsSeparateWithTheirNesting() {
        var flat = new PdfCore.PdfListStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Zero };
        var nested = new PdfCore.PdfListStyle { LeftIndent = global::OfficeIMO.OfficeLength.Points(18), SpacingBefore = global::OfficeIMO.OfficeLength.Zero };
        var document = PdfCore.PdfDocument.Create();
        document.AddList(new[] { "Reduce backlog" }, style: flat);
        document.AddList(new[] { "Archive exports", "Notify owners" }, style: nested);
        document.AddList(new[] { "Expand coverage" }, style: flat);
        document.AddParagraph("Next steps follow.");
        document.AddNumberedList(new[] { "Collect requirements", "Agree on scope" });
        document.AddParagraph("Follow-up actions.");
        document.AddNumberedList(new[] { "Add alert", "Rehearse" }, startNumber: 3);
        byte[] pdf = document.ToBytes();

        using OfficeWordDocument imported = LoadSemanticPdf(pdf).ConvertToWordWithEvidence().Value;
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(imported.ToBytes()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Body body = GetPdfSemanticBody(package);
        Numbering numbering = package.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;

        (int NumberingId, int Level, NumberFormatValues Format, int Start) Entry(string text) {
            Paragraph paragraph = Assert.Single(body.Elements<Paragraph>(), item => ReadParagraphText(item) == text);
            NumberingProperties properties = paragraph.ParagraphProperties?.NumberingProperties
                ?? throw new Xunit.Sdk.XunitException(text + " is not a list item.");
            int numId = properties.NumberingId!.Val!.Value;
            int level = properties.NumberingLevelReference!.Val!.Value;
            NumberingInstance instance = numbering.Elements<NumberingInstance>().Single(item => item.NumberID!.Value == numId);
            AbstractNum definition = numbering.Elements<AbstractNum>().Single(item => item.AbstractNumberId!.Value == instance.AbstractNumId!.Val!.Value);
            Level levelDefinition = definition.Elements<Level>().Single(item => item.LevelIndex!.Value == level);
            return (numId, level, levelDefinition.NumberingFormat!.Val!.Value, levelDefinition.StartNumberingValue?.Val?.Value ?? 1);
        }

        var parent = Entry("Reduce backlog");
        var child = Entry("Archive exports");
        Assert.Equal(NumberFormatValues.Bullet, parent.Format);
        Assert.Equal(0, parent.Level);
        Assert.Equal(1, child.Level);
        Assert.Equal(parent.NumberingId, child.NumberingId);
        Assert.Equal(NumberFormatValues.Bullet, child.Format);
        Assert.Equal(parent.NumberingId, Entry("Expand coverage").NumberingId);

        var firstNumbered = Entry("Collect requirements");
        var secondNumbered = Entry("Add alert");
        Assert.Equal(NumberFormatValues.Decimal, firstNumbered.Format);
        Assert.Equal(1, firstNumbered.Start);
        Assert.NotEqual(parent.NumberingId, firstNumbered.NumberingId);
        // A separate list does not count on from the first one, and keeps its own start number.
        Assert.NotEqual(firstNumbered.NumberingId, secondNumbered.NumberingId);
        Assert.Equal(NumberFormatValues.Decimal, secondNumbered.Format);
        Assert.Equal(3, secondNumbered.Start);
    }

    [Fact]
    public void PdfSemanticImport_LinksCoverOnlyTheirOwnWordsWithoutRepeatingTheLine() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => {
                paragraph.AddText("See the ");
                paragraph.AddLink("project page", "https://example.com/project");
                paragraph.AddText(" and the ");
                paragraph.AddLink("archive", "https://example.org/archive");
                paragraph.AddText(" for details.");
            })
            .ToBytes();

        using OfficeWordDocument imported = LoadSemanticPdf(pdf).ConvertToWordWithEvidence().Value;
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(imported.ToBytes()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Body body = GetPdfSemanticBody(package);

        Paragraph paragraph = Assert.Single(body.Elements<Paragraph>(), item => ReadParagraphText(item).Length > 0);
        Assert.Equal("See the project page and the archive for details.", ReadParagraphText(paragraph));
        List<Hyperlink> links = paragraph.Elements<Hyperlink>().ToList();
        Assert.Equal(new[] { "project page", "archive" }, links.Select(ReadHyperlinkText));
        Dictionary<string, Uri> targets = package.MainDocumentPart!.HyperlinkRelationships.ToDictionary(item => item.Id, item => item.Uri);
        Assert.Equal(new Uri("https://example.com/project"), targets[links[0].Id!.Value!]);
        Assert.Equal(new Uri("https://example.org/archive"), targets[links[1].Id!.Value!]);
    }

    [Fact]
    public void PdfSemanticImport_KeepsLinksThatCoverPartOfAWord() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => {
                paragraph.AddText("Read Open");
                paragraph.AddLink("AI", "https://example.com/ai");
                paragraph.AddText(" notes and ");
                paragraph.AddLink("Alpha", "https://example.com/alpha");
                paragraph.AddText("/");
                paragraph.AddLink("Beta", "https://example.com/beta");
                paragraph.AddText(" today.");
            })
            .ToBytes();

        using OfficeWordDocument imported = LoadSemanticPdf(pdf).ConvertToWordWithEvidence().Value;
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(imported.ToBytes()), false);
        Body body = GetPdfSemanticBody(package);

        Paragraph paragraph = Assert.Single(body.Elements<Paragraph>(), item => ReadParagraphText(item).Length > 0);
        Assert.Equal("Read OpenAI notes and Alpha/Beta today.", ReadParagraphText(paragraph));
        Dictionary<string, Uri> targets = package.MainDocumentPart!.HyperlinkRelationships.ToDictionary(item => item.Id, item => item.Uri);
        Assert.Equal(
            new[] { ("AI", "https://example.com/ai"), ("Alpha", "https://example.com/alpha"), ("Beta", "https://example.com/beta") },
            paragraph.Elements<Hyperlink>().Select(link => (ReadHyperlinkText(link), targets[link.Id!.Value!].OriginalString)));
    }
}

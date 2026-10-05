using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LegacyDoc_NoteReferencesRetainCharacterStyleTypography(bool endnote, bool directOverride) {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("StyledNoteMarker");
        if (endnote) paragraph.AddEndNote("Styled endnote"); else paragraph.AddFootNote("Styled footnote");
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style inherited = styles.Elements<Style>().Single(style => style.StyleId == "DefaultParagraphFont");
        inherited.StyleRunProperties = new StyleRunProperties(
            new RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
            new Color { Val = "3355CC" });
        Style noteStyle = styles.Elements<Style>().Single(style => style.StyleId == (endnote ? "EndnoteReference" : "FootnoteReference"));
        noteStyle.StyleRunProperties ??= new StyleRunProperties();
        noteStyle.StyleRunProperties.Append(new FontSize { Val = "28" }, new FontSizeComplexScript { Val = "28" });
        Run reference = paragraph._paragraph.Descendants<Run>().Single(run => endnote
            ? run.Elements<EndnoteReference>().Any() : run.Elements<FootnoteReference>().Any());
        if (directOverride) {
            reference.RunProperties!.Append(
                new FontSize { Val = "36" }, new FontSizeComplexScript { Val = "36" },
                new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" },
                new Color { Val = "990000" });
        }
        using var reopened = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        Run imported = reopened._document!.Body!.Descendants<Run>().Single(run => endnote
            ? run.Elements<EndnoteReference>().Any() : run.Elements<FootnoteReference>().Any());
        Assert.Equal(directOverride ? "36" : "28", imported.RunProperties?.FontSize?.Val?.Value);
        Assert.Equal(directOverride ? "Times New Roman" : "Arial", imported.RunProperties?.RunFonts?.Ascii?.Value);
        Assert.Equal(directOverride ? "990000" : "3355CC", imported.RunProperties?.Color?.Val?.Value);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LegacyDoc_NoteReferencesRetainRevisionIdentityAndAcceptance(bool endnote, bool inserted) {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("RevisedNoteMarker");
        if (endnote) paragraph.AddEndNote("Revised endnote"); else paragraph.AddFootNote("Revised footnote");
        Run reference = paragraph._paragraph.Descendants<Run>().Single(run => endnote
            ? run.Elements<EndnoteReference>().Any() : run.Elements<FootnoteReference>().Any());
        reference.RunProperties!.Append(new FontSize { Val = "28" }, new FontSizeComplexScript { Val = "28" });
        var date = new DateTime(2026, 10, 3, 10, 20, 0, DateTimeKind.Utc);
        OpenXmlCompositeElement revision = inserted
            ? new InsertedRun { Id = "123", Author = "Reference author", Date = date }
            : new DeletedRun { Id = "123", Author = "Reference author", Date = date };
        reference.InsertBeforeSelf(revision);
        reference.Remove();
        revision.Append(reference);
        using var reopened = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        Run imported = reopened._document!.Body!.Descendants<Run>().Single(run => endnote
            ? run.Elements<EndnoteReference>().Any() : run.Elements<FootnoteReference>().Any());
        Assert.Equal("28", imported.RunProperties?.FontSize?.Val?.Value);
        if (inserted) {
            InsertedRun wrapper = Assert.IsType<InsertedRun>(imported.Parent);
            Assert.Equal("Reference author", wrapper.Author?.Value);
            Assert.Equal(date, wrapper.Date?.Value);
        } else {
            DeletedRun wrapper = Assert.IsType<DeletedRun>(imported.Parent);
            Assert.Equal("Reference author", wrapper.Author?.Value);
            Assert.Equal(date, wrapper.Date?.Value);
        }
        reopened.AcceptRevisions();
        Assert.Equal(inserted ? 1 : 0, reopened._document.Body.Descendants<Run>().Count(run => endnote
            ? run.Elements<EndnoteReference>().Any() : run.Elements<FootnoteReference>().Any()));
    }
}

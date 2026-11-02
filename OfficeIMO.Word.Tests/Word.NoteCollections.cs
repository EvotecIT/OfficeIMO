using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordNoteCollectionTests {
    [Fact]
    public void NoteCollectionsPreserveBodyAndNestedTableOrderAcrossSectionsAfterReload() {
        using WordDocument document = WordDocument.Create();
        AddNotePair(document.AddParagraph(), "first body");
        WordTable table = document.AddTable(1, 2);
        WordTableCell firstCell = table.Rows[0].Cells[0];
        AddNotePair(firstCell.Paragraphs[0], "first cell");
        WordTable nested = firstCell.AddTable(1, 1);
        AddNotePair(nested.Rows[0].Cells[0].Paragraphs[0], "nested cell");
        AddNotePair(table.Rows[0].Cells[1].Paragraphs[0], "second cell");
        AddNotePair(document.AddParagraph(), "last body");
        document.AddSection();
        AddNotePair(document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0], "second section cell");
        AddNotePair(document.AddParagraph(), "second section body");

        AssertMixedNoteOrder(document);
        Assert.Empty(document.ValidateDocument());

        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes()));
        AssertMixedNoteOrder(restored);
        Assert.Empty(restored.ValidateDocument());
    }

    [Fact]
    public void CollectedTableNoteHandlesEditAndRemoveTheirOriginalOwnersAfterReload() {
        using WordDocument document = WordDocument.Create();
        AddNotePair(document.AddParagraph(), "body");
        AddNotePair(document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0], "table");
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes()));

        WordFootNote footnote = Assert.Single(restored.FootNotes, note =>
            NoteText(note.Paragraphs) == "table footnote");
        WordEndNote endnote = Assert.Single(restored.EndNotes, note =>
            NoteText(note.Paragraphs) == "table endnote");
        WordParagraph footnoteOwner = Assert.IsType<WordParagraph>(footnote.ParentParagraph);
        WordParagraph endnoteOwner = Assert.IsType<WordParagraph>(endnote.ParentParagraph);
        Assert.IsType<WordTableCell>(footnoteOwner.Parent);
        Assert.IsType<WordTableCell>(endnoteOwner.Parent);
        footnoteOwner.Text = "edited footnote owner";
        endnoteOwner.Text = "edited endnote owner";
        Assert.NotNull(footnote.Paragraphs);
        Assert.NotNull(endnote.Paragraphs);
        footnote.Paragraphs!.Last().Text = "edited footnote";
        endnote.Paragraphs!.Last().Text = "edited endnote";

        using WordDocument edited = WordDocument.Load(new MemoryStream(restored.ToBytes()));
        Assert.Equal(new[] { "body footnote", "edited footnote" },
            edited.FootNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Equal(new[] { "body endnote", "edited endnote" },
            edited.EndNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Equal("edited footnote owner", edited.FootNotes[1].ParentParagraph!.Text);
        Assert.Equal("edited endnote owner", edited.EndNotes[1].ParentParagraph!.Text);

        edited.FootNotes[1].Remove();
        edited.EndNotes[1].Remove();
        Assert.Equal("body footnote", NoteText(Assert.Single(edited.FootNotes).Paragraphs));
        Assert.Equal("body endnote", NoteText(Assert.Single(edited.EndNotes).Paragraphs));
        Assert.Contains(edited.Tables[0].Rows[0].Cells[0].Paragraphs,
            paragraph => paragraph.Text == "edited footnote owner");
        Assert.Contains(edited.Tables[0].Rows[0].Cells[0].Paragraphs,
            paragraph => paragraph.Text == "edited endnote owner");
        Assert.Empty(edited.ValidateDocument());

        using WordDocument removed = WordDocument.Load(new MemoryStream(edited.ToBytes()));
        Assert.Equal("body footnote", NoteText(Assert.Single(removed.FootNotes).Paragraphs));
        Assert.Equal("body endnote", NoteText(Assert.Single(removed.EndNotes).Paragraphs));
        Assert.DoesNotContain(removed.Tables[0].Rows[0].Cells[0].Paragraphs,
            paragraph => paragraph.IsFootNote || paragraph.IsEndNote);
        Assert.Empty(removed.ValidateDocument());
    }

    [Fact]
    public void FootnoteDefinitionsRemainAvailableOnlyWhenNoBodyReferenceExists() {
        using WordDocument document = WordDocument.Create();
        WordParagraph removedReference = document.AddParagraph("orphan owner").AddFootNote("orphan definition").Paragraph;
        removedReference.Remove();
        Assert.Empty(document.Sections[0].FootNotes);
        Assert.Equal("orphan definition", NoteText(Assert.Single(document.FootNotes).Paragraphs));

        WordParagraph liveReference = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0]
            .AddText("live owner").AddFootNote("live definition").Paragraph;
        Assert.Equal("live definition", NoteText(Assert.Single(document.FootNotes).Paragraphs));
        Assert.Single(document.Sections[0].FootNotes);
        liveReference.Remove();

        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(restored.Sections[0].FootNotes);
        Assert.Equal(new[] { "orphan definition", "live definition" },
            restored.FootNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Empty(restored.EndNotes);
        Assert.Empty(restored.ValidateDocument());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void RevisionOnlyFootnoteReferencesDoNotBecomeSyntheticLiveNotes(bool movedFrom, bool orphanDefinition) {
        byte[] source = CreateRevisionOnlyFootnote(movedFrom, orphanDefinition);
        var expected = ReadRevisionFootnoteSource(source);
        Assert.Single(expected.References);
        Assert.Equal(orphanDefinition ? 2 : 1, expected.Definitions.Length);

        using WordDocument document = WordDocument.Load(new MemoryStream(source));
        Assert.Empty(document.Sections[0].FootNotes);
        Assert.Empty(document.FootNotes);
        Assert.Empty(document.ValidateDocument());
        byte[] saved = document.ToBytes();
        var actual = ReadRevisionFootnoteSource(saved);
        Assert.Equal(expected.References, actual.References);
        Assert.Equal(expected.Definitions, actual.Definitions);
        Assert.Equal(expected.Revisions, actual.Revisions);

        using WordDocument restored = WordDocument.Load(new MemoryStream(saved));
        Assert.Empty(restored.Sections[0].FootNotes);
        Assert.Empty(restored.FootNotes);
        Assert.Empty(restored.ValidateDocument());
        var rewritten = ReadRevisionFootnoteSource(restored.ToBytes());
        Assert.Equal(expected.References, rewritten.References);
        Assert.Equal(expected.Definitions, rewritten.Definitions);
        Assert.Equal(expected.Revisions, rewritten.Revisions);
    }

    private static byte[] CreateRevisionOnlyFootnote(bool movedFrom, bool orphanDefinition) {
        using WordDocument document = WordDocument.Create();
        document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0]
            .AddText("revision owner").AddFootNote("revision definition");
        if (orphanDefinition) {
            document.AddParagraph("orphan owner").AddFootNote("orphan definition").Paragraph.Remove();
        }

        using var imported = new MemoryStream();
        byte[] original = document.ToBytes();
        imported.Write(original, 0, original.Length);
        imported.Position = 0;
        using (WordprocessingDocument package = WordprocessingDocument.Open(imported, true)) {
            Run referenceRun = Assert.Single(package.MainDocumentPart!.Document.Body!.Descendants<Run>(),
                run => run.Elements<FootnoteReference>().Any());
            OpenXmlCompositeElement revision = movedFrom
                ? new MoveFromRun { Id = "1", Author = "OfficeIMO Tests" }
                : new DeletedRun { Id = "1", Author = "OfficeIMO Tests" };
            referenceRun.InsertBeforeSelf(revision);
            referenceRun.Remove();
            revision.Append(referenceRun);
            package.MainDocumentPart.Document.Save();
        }
        return imported.ToArray();
    }

    private static (string[] References, string[] Definitions, string[] Revisions) ReadRevisionFootnoteSource(byte[] bytes) {
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        Body body = package.MainDocumentPart!.Document.Body!;
        return (
            body.Descendants<FootnoteReference>().Select(reference => reference.OuterXml).ToArray(),
            package.MainDocumentPart.FootnotesPart!.Footnotes!.Elements<Footnote>()
                .Where(note => note.Type == null).Select(note => note.OuterXml).ToArray(),
            body.Descendants().Where(element => element is DeletedRun || element is MoveFromRun)
                .Select(element => element.OuterXml).ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SeveralSameKindNotesInOneImportedRunKeepIndependentIdentityAndSiblingContent(bool endnotes) {
        byte[] source = CreateSharedNoteRun(endnotes);
        using WordDocument document = WordDocument.Load(new MemoryStream(source));
        Assert.Empty(document.ValidateDocument());
        long?[] originalIds = endnotes
            ? document.EndNotes.Select(note => note.ReferenceId).ToArray()
            : document.FootNotes.Select(note => note.ReferenceId).ToArray();
        long? companionId = endnotes
            ? Assert.Single(document.FootNotes).ReferenceId
            : Assert.Single(document.EndNotes).ReferenceId;
        Assert.Equal(3, originalIds.Length);
        Assert.Equal(3, originalIds.Distinct().Count());
        Assert.Equal(new[] { "first note", "second note", "third note" }, endnotes
            ? document.EndNotes.Select(note => NoteText(note.Paragraphs))
            : document.FootNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Equal(originalIds, endnotes
            ? document.Sections[0].EndNotes.Select(note => note.ReferenceId)
            : document.Sections[0].FootNotes.Select(note => note.ReferenceId));
        AssertSharedNoteSource(source, endnotes, originalIds, companionId);

        if (endnotes) {
            document.EndNotes[1].Paragraphs!.Last().Text = "edited second note";
        } else {
            document.FootNotes[1].Paragraphs!.Last().Text = "edited second note";
        }

        using WordDocument edited = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(new[] { "first note", "edited second note", "third note" }, endnotes
            ? edited.EndNotes.Select(note => NoteText(note.Paragraphs))
            : edited.FootNotes.Select(note => NoteText(note.Paragraphs)));
        AssertSharedNoteSource(edited.ToBytes(), endnotes, originalIds, companionId);
        if (endnotes) {
            edited.EndNotes[1].Remove();
        } else {
            edited.FootNotes[1].Remove();
        }
        Assert.Equal(new[] { "first note", "third note" }, endnotes
            ? edited.EndNotes.Select(note => NoteText(note.Paragraphs))
            : edited.FootNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Equal("companion note", endnotes
            ? NoteText(Assert.Single(edited.FootNotes).Paragraphs)
            : NoteText(Assert.Single(edited.EndNotes).Paragraphs));
        Assert.Empty(edited.ValidateDocument());

        byte[] removedBytes = edited.ToBytes();
        AssertSharedNoteSource(removedBytes, endnotes, new[] { originalIds[0], originalIds[2] }, companionId);
        using WordDocument removed = WordDocument.Load(new MemoryStream(removedBytes));
        Assert.Equal(new[] { originalIds[0], originalIds[2] }, endnotes
            ? removed.EndNotes.Select(note => note.ReferenceId)
            : removed.FootNotes.Select(note => note.ReferenceId));
        Assert.Equal(new[] { "first note", "third note" }, endnotes
            ? removed.EndNotes.Select(note => NoteText(note.Paragraphs))
            : removed.FootNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Empty(removed.ValidateDocument());
    }

    private static byte[] CreateSharedNoteRun(bool endnotes) {
        using WordDocument document = WordDocument.Create();
        WordParagraph owner = document.AddParagraph("owner");
        foreach (string text in new[] { "first note", "second note", "third note" }) {
            if (endnotes) {
                owner.AddEndNote(text);
            } else {
                owner.AddFootNote(text);
            }
            if (text == "first note") {
                owner.AddText(" shared text ");
            }
        }
        if (endnotes) {
            owner.AddFootNote("companion note");
        } else {
            owner.AddEndNote("companion note");
        }

        using var imported = new MemoryStream();
        byte[] original = document.ToBytes();
        imported.Write(original, 0, original.Length);
        imported.Position = 0;
        using (WordprocessingDocument package = WordprocessingDocument.Open(imported, true)) {
            Paragraph paragraph = Assert.Single(package.MainDocumentPart!.Document.Body!.Elements<Paragraph>());
            var shared = new Run(new RunProperties(new Bold()));
            foreach (Run run in paragraph.Elements<Run>().Skip(1).ToList()) {
                foreach (OpenXmlElement child in run.ChildElements.Where(child => child is not RunProperties).ToList()) {
                    child.Remove();
                    shared.Append(child);
                }
                run.Remove();
            }
            paragraph.Append(shared);
            package.MainDocumentPart.Document.Save();
        }
        return imported.ToArray();
    }

    private static void AssertSharedNoteSource(byte[] bytes, bool endnotes, IReadOnlyList<long?> expectedIds,
        long? companionId) {
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        Paragraph paragraph = Assert.Single(package.MainDocumentPart!.Document.Body!.Elements<Paragraph>());
        Assert.Equal("owner shared text ", paragraph.InnerText);
        Run shared = paragraph.Elements<Run>().Last();
        Assert.NotNull(shared.RunProperties?.Bold);
        Assert.Equal(" shared text ", string.Concat(shared.Elements<Text>().Select(text => text.Text)));
        Assert.Equal(expectedIds, endnotes
            ? shared.Elements<EndnoteReference>().Select(reference => reference.Id?.Value)
            : shared.Elements<FootnoteReference>().Select(reference => reference.Id?.Value));
        Assert.Equal(expectedIds, endnotes
            ? package.MainDocumentPart.EndnotesPart!.Endnotes!.Elements<Endnote>()
                .Where(note => note.Type == null).Select(note => note.Id?.Value)
            : package.MainDocumentPart.FootnotesPart!.Footnotes!.Elements<Footnote>()
                .Where(note => note.Type == null).Select(note => note.Id?.Value));
        Assert.Equal(companionId, endnotes
            ? Assert.Single(shared.Elements<FootnoteReference>()).Id?.Value
            : Assert.Single(shared.Elements<EndnoteReference>()).Id?.Value);
        Assert.Equal(2, paragraph.Elements<Run>().Count());
    }

    private static void AddNotePair(WordParagraph paragraph, string name) {
        paragraph.Text = name;
        paragraph.AddFootNote(name + " footnote").Paragraph
            .AddText(name + " endnote owner").AddEndNote(name + " endnote");
    }

    private static void AssertMixedNoteOrder(WordDocument document) {
        string[] firstSection = { "first body", "first cell", "nested cell", "second cell", "last body" };
        string[] secondSection = { "second section cell", "second section body" };
        Assert.Equal(2, document.Sections.Count);
        Assert.Equal(firstSection.Select(name => name + " footnote"),
            document.Sections[0].FootNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Equal(firstSection.Select(name => name + " endnote"),
            document.Sections[0].EndNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Equal(secondSection.Select(name => name + " footnote"),
            document.Sections[1].FootNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Equal(secondSection.Select(name => name + " endnote"),
            document.Sections[1].EndNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Equal(firstSection.Concat(secondSection).Select(name => name + " footnote"),
            document.FootNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Equal(firstSection.Concat(secondSection).Select(name => name + " endnote"),
            document.EndNotes.Select(note => NoteText(note.Paragraphs)));
        Assert.Equal(firstSection.Concat(secondSection),
            document.FootNotes.Select(note => note.ParentParagraph!.Text));
        Assert.Equal(firstSection.Concat(secondSection).Select(name => name + " endnote owner"),
            document.EndNotes.Select(note => note.ParentParagraph!.Text));
    }

    private static string NoteText(List<WordParagraph>? paragraphs) {
        Assert.NotNull(paragraphs);
        return string.Concat(paragraphs!.Select(paragraph => paragraph.Text));
    }
}

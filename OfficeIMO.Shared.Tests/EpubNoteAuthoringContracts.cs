using OfficeIMO.Epub;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubNoteAuthoringContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly XNamespace Ops = "http://www.idpf.org/2007/ops";

    [Theory]
    [InlineData(false, EpubNoteKind.Footnote)]
    [InlineData(true, EpubNoteKind.Footnote)]
    [InlineData(false, EpubNoteKind.Endnote)]
    public void NotesRoundTripWithForwardAndReturnLinks(bool sameDocument, EpubNoteKind kind) {
        var book = Book(sameDocument, kind);
        var options = Options(sameDocument, kind);
        book.AddNote(options);
        var loaded = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        XDocument source = loaded.GetContentXml("source");
        XDocument notes = loaded.GetContentXml(options.NotesManifestId);
        XElement marker = ById(source, "ref");
        XElement note = ById(notes, "note-1");
        Assert.Equal("1", marker.Value);
        Assert.Equal("retained", (string?)marker.Attribute("class"));
        Assert.Equal("doc-noteref", (string?)marker.Attribute("role"));
        Assert.Equal("noteref", (string?)marker.Attribute(Ops + "type"));
        EpubReference forward = EpubReference.Resolve("EPUB/text/source.xhtml", marker.Attribute("href")!.Value);
        Assert.Equal(sameDocument ? "EPUB/text/source.xhtml" : "EPUB/back/notes.xhtml", forward.ContainerPath);
        Assert.Equal("note-1", forward.Fragment);
        XElement backlink = note.Descendants(Html + "a").Single();
        EpubReference back = EpubReference.Resolve(forward.ContainerPath!, backlink.Attribute("href")!.Value);
        Assert.Equal("EPUB/text/source.xhtml", back.ContainerPath);
        Assert.Equal("ref", back.Fragment);
        Assert.Equal("Return to text", backlink.Value);
        Assert.Equal("doc-backlink", (string?)backlink.Attribute("role"));
        Assert.Equal("Preserved", source.Descendants(Html + "meta").Single().Attribute("content")!.Value);
        if (kind == EpubNoteKind.Endnote) {
            Assert.Equal(Html + "li", note.Name);
            Assert.Null(note.Attribute("role"));
            Assert.Equal("doc-endnotes", (string?)note.Ancestors(Html + "section").First().Attribute("role"));
        } else {
            Assert.Equal(Html + "aside", note.Name);
            Assert.Equal("doc-footnote", (string?)note.Attribute("role"));
        }
    }

    [Fact]
    public void GeneratedNoteLinksRespectBothHtmlBases() {
        var book = Book(false, EpubNoteKind.Footnote);
        foreach (string id in new[] { "source", "notes" }) {
            XDocument content = book.GetContentXml(id);
            content.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", "../assets/")));
            book.SetContentXml(id, content);
        }
        book.AddNote(Options(false, EpubNoteKind.Footnote));
        var loaded = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var reference = EpubReference.Resolve("EPUB/text/source.xhtml", "../assets/", ById(loaded.GetContentXml("source"), "ref").Attribute("href")!.Value);
        Assert.Equal("EPUB/back/notes.xhtml", reference.ContainerPath);
        var backlink = loaded.GetContentXml("notes").Descendants(Html + "a").Single();
        Assert.Equal("EPUB/text/source.xhtml", EpubReference.Resolve("EPUB/back/notes.xhtml", "../assets/", backlink.Attribute("href")!.Value).ContainerPath);
    }

    [Theory]
    [InlineData("duplicate")]
    [InlineData("script")]
    [InlineData("missing-idref")]
    [InlineData("linked-marker")]
    public void InvalidSecondDocumentOrExistingLinkCannotPartiallyEditTheSource(string failure) {
        var book = Book(false, EpubNoteKind.Footnote);
        var options = Options(false, EpubNoteKind.Footnote);
        if (failure == "linked-marker") {
            XDocument source = book.GetContentXml("source");
            ById(source, "ref").SetAttributeValue("href", "#ref");
            book.SetContentXml("source", source);
        } else options.BodyXhtml = failure == "duplicate" ? "<p id='note-1'>Collision</p>" :
            failure == "script" ? "<script>bad()</script>" : "<p aria-describedby='missing'>Text</p>";
        byte[] before = book.Write().Bytes;
        if (failure == "script") Assert.Throws<NotSupportedException>(() => book.AddNote(options));
        else Assert.Throws<InvalidDataException>(() => book.AddNote(options));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void CombinedRetentionLimitAndCancellationLeaveBothDocumentsUnchanged() {
        var original = Book(false, EpubNoteKind.Footnote);
        byte[] input = original.Write().Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var book = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions {
            MaxExpandedBytes = zip.Entries.Sum(entry => entry.Length) + 128
        });
        Assert.Throws<InvalidDataException>(() => book.AddNote(Options(false, EpubNoteKind.Footnote)));
        Assert.Equal(input, book.Write().Bytes);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => original.AddNote(Options(false, EpubNoteKind.Footnote), cancellation.Token));
        Assert.Equal(input, original.Write().Bytes);
    }

    private static XElement ById(XDocument document, string id) => document.Descendants().Single(element => (string?)element.Attribute("id") == id);
    private static EpubNoteOptions Options(bool sameDocument, EpubNoteKind kind) => new EpubNoteOptions {
        SourceManifestId = "source", ReferenceId = "ref", NotesManifestId = sameDocument ? "source" : "notes",
        ContainerId = "notes-container", NoteId = "note-1", BodyXhtml = "<p>A <em>reviewed</em> note.</p>", BacklinkText = "Return to text", Kind = kind
    };
    private static EpubPublication Book(bool sameDocument, EpubNoteKind kind) {
        var book = EpubPublication.Create("Notes", "en");
        string notes = kind == EpubNoteKind.Footnote ? "<div id='notes-container'/>" : "<section><h1>Notes</h1><ol id='notes-container'/></section>";
        book.AddChapter("source", "EPUB/text/source.xhtml", "Chapter", "<h1>Chapter</h1><p>A statement <a id='ref' class='retained'>1</a>.</p>" + (sameDocument ? notes : string.Empty));
        if (!sameDocument) book.AddChapter("notes", "EPUB/back/notes.xhtml", "Notes", notes);
        XDocument content = book.GetContentXml("source");
        content.Root!.Element(Html + "head")!.Add(new XElement(Html + "meta", new XAttribute("name", "description"), new XAttribute("content", "Preserved")));
        book.SetContentXml("source", content);
        return book;
    }
}

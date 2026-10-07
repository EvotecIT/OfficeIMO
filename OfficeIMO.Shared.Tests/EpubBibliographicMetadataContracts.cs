using OfficeIMO.Epub;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubBibliographicMetadataContracts {
    [Fact]
    public void TitleVariantsAndPrimaryUpdatesRetainRefinementsAndOptionalFields() {
        var book = Book();
        book.SetPrimaryTitle("main-title", new EpubTitleMetadata { Text = "The Book", Language = "en", FileAs = "Book, The", DisplaySequence = 1 });
        book.AddMetadataProperty("alternate-script", "Das Buch", "#main-title");
        book.AddTitle("subtitle", new EpubTitleMetadata { Text = "An introduction", Kind = EpubTitleKind.Subtitle, DisplaySequence = 2 });
        book.SetPrimaryTitle("main-title", new EpubTitleMetadata { Text = "The Revised Book" });
        var read = Reopen(book);
        Assert.Equal("The Revised Book", book.Title);
        var primary = Assert.Single(read.Metadata, entry => entry.Id == "main-title");
        Assert.Equal("en", primary.Language);
        Assert.Equal("Book, The", primary.FileAs);
        Assert.Contains(read.Metadata, entry => entry.Refines == "#main-title" && entry.Property == "alternate-script" && entry.Value == "Das Buch");
        Assert.Equal("1", Assert.Single(read.Metadata, entry => entry.Refines == "#main-title" && entry.Property == "display-seq").Value);
        Assert.Equal("main", Assert.Single(read.Metadata, entry => entry.Refines == "#main-title" && entry.Property == "title-type").Value);
        Assert.Contains(read.Metadata, entry => entry.Refines == "#subtitle" && entry.Property == "title-type" && entry.Value == "subtitle");
        byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.SetPrimaryTitle("replacement-id", new EpubTitleMetadata { Text = "Broken reference" }));
        Assert.Throws<ArgumentException>(() => book.AddTitle("chapter", new EpubTitleMetadata { Text = "Duplicate id" }));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData(EpubIdentifierKind.Isbn13, "978-0-306-40615-7", "urn:isbn:9780306406157", "15")]
    [InlineData(EpubIdentifierKind.Isbn10, "urn:isbn:0-306-40615-2", "urn:isbn:0306406152", "02")]
    [InlineData(EpubIdentifierKind.Isbn10, "0 8044 2957 x", "urn:isbn:080442957X", "02")]
    [InlineData(EpubIdentifierKind.Doi, "10.1000/182", "10.1000/182", "06")]
    [InlineData(EpubIdentifierKind.Unspecified, "urn:example:catalog:1", "urn:example:catalog:1", null)]
    public void IdentifierRecordsNormalizeIsbnsWithoutReplacingThePackageIdentity(EpubIdentifierKind kind, string input, string expected, string? code) {
        var book = Book();
        string identity = book.Identifier;
        book.AddIdentifier("catalog-id", new EpubIdentifierMetadata { Kind = kind, Value = input });
        var read = Reopen(book);
        Assert.Equal(expected, Assert.Single(read.Metadata, entry => entry.Id == "catalog-id").Value);
        Assert.Equal(identity, book.Identifier);
        var refinement = read.Metadata.SingleOrDefault(entry => entry.Refines == "#catalog-id" && entry.Property == "identifier-type");
        if (code == null) Assert.Null(refinement);
        else { Assert.Equal(code, refinement!.Value); Assert.Equal("onix:codelist5", refinement.Scheme); }
    }

    [Theory]
    [InlineData(EpubIdentifierKind.Isbn13, "9780306406158")]
    [InlineData(EpubIdentifierKind.Isbn13, "4006381333931")]
    [InlineData(EpubIdentifierKind.Isbn10, "0306406153")]
    [InlineData(EpubIdentifierKind.Isbn10, "X306406152")]
    [InlineData(EpubIdentifierKind.Doi, "https://doi.org/10.1000/182")]
    [InlineData(EpubIdentifierKind.Doi, "10.1000/")]
    public void InvalidIdentifierRecordsDoNotMutateThePackage(EpubIdentifierKind kind, string value) {
        var book = Book();
        byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.AddIdentifier("bad-id", new EpubIdentifierMetadata { Kind = kind, Value = value }));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void SubjectsAndPublicationDetailsPreserveOtherValuesAndSerializeCalendarDates() {
        var book = Book();
        book.AddDublinCoreMetadata("rights", "Prior rights", "rights");
        book.AddMetadataProperty("alternate-script", "Retained translation", "#rights");
        book.AddDublinCoreMetadata("publisher", "Primary publisher");
        book.AddDublinCoreMetadata("publisher", "Co-publisher");
        book.AddSubject("subject", new EpubSubjectMetadata { Text = "FICTION / General", Authority = "BISAC", Code = "FIC000000", Language = "en" });
        book.SetPublicationDetails(new EpubPublicationDetails {
            Publisher = "Updated publisher", Description = "An editorial description.", Rights = "Revised rights",
            PublicationDate = new DateTime(2026, 10, 5, 23, 59, 0, DateTimeKind.Local)
        });
        book.SetPublicationDetails(new EpubPublicationDetails { Description = "Revised description." });
        var metadata = Reopen(book).Metadata;
        Assert.Equal(new[] { "Updated publisher", "Co-publisher" }, metadata.Where(entry => entry.Name == "publisher").Select(entry => entry.Value));
        Assert.Contains(metadata, entry => entry.Name == "date" && entry.Value == "2026-10-05");
        Assert.Contains(metadata, entry => entry.Id == "rights" && entry.Value == "Revised rights");
        Assert.Contains(metadata, entry => entry.Refines == "#rights" && entry.Value == "Retained translation");
        Assert.Contains(metadata, entry => entry.Refines == "#subject" && entry.Property == "authority" && entry.Value == "BISAC");
        Assert.Contains(metadata, entry => entry.Refines == "#subject" && entry.Property == "term" && entry.Value == "FIC000000");
        byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.AddSubject("bad", new EpubSubjectMetadata { Text = "Subject", Code = "Missing authority" }));
        Assert.Throws<ArgumentException>(() => book.SetPublicationDetails(new EpubPublicationDetails { Publisher = "New", Rights = " " }));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void BibliographicEditsRespectCombinedBudgetsAndVocabularyBindings() {
        byte[] input = Book().Write().Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxExpandedBytes = zip.Entries.Sum(entry => entry.Length) + 40 });
        Assert.Throws<InvalidDataException>(() => limited.SetPrimaryTitle("main-title", new EpubTitleMetadata { Text = "Replacement title", FileAs = "Title, replacement", DisplaySequence = 1 }));
        Assert.Throws<InvalidDataException>(() => limited.SetPublicationDetails(new EpubPublicationDetails { Publisher = "Publisher", Rights = new string('r', 100) }));
        Assert.Throws<InvalidDataException>(() => limited.AddSubject("subject", new EpubSubjectMetadata { Text = "Fiction", Authority = "BISAC", Code = "FIC000000" }));
        Assert.Equal(input, limited.Write().Bytes);
        var book = Book();
        book.DeclareVocabularyPrefix("onix", "https://example.org/not-onix/");
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidOperationException>(() => book.AddIdentifier("isbn", new EpubIdentifierMetadata { Value = "9780306406157", Kind = EpubIdentifierKind.Isbn13 }));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubDocument Reopen(EpubPublication book) => EpubDocument.Load(new MemoryStream(book.Write().Bytes));
    private static EpubPublication Book() {
        var book = EpubPublication.Create("Bibliographic metadata", "en");
        book.AddChapter("chapter", "EPUB/chapter.xhtml", "Chapter", "<h1>Chapter</h1><p>Text.</p>");
        return book;
    }
}

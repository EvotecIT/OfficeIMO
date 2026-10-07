using OfficeIMO.Epub;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubPublishingMetadataContracts {
    [Fact]
    public void ImportedRefinementsResolveOnlyPackageVocabularyAndRetainLegacyPrecedence() {
        var book = Book();
        XDocument package = book.GetPackageXml();
        XNamespace opf = "http://www.idpf.org/2007/opf";
        XElement metadata = package.Root!.Element(opf + "metadata")!;
        metadata.Add(XElement.Parse("""
            <dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/" id="author-name">Writer</dc:creator>
            """));
        metadata.Add(XElement.Parse("""
            <meta xmlns="https://example.org/foreign" property="role" refines="#author-name">wrong</meta>
            """));
        metadata.Add(XElement.Parse("""
            <meta xmlns="http://www.idpf.org/2007/opf" property="role" refines="#author%2Dname" scheme="marc:relators">aut</meta>
            """));
        metadata.Add(XElement.Parse("""
            <meta xmlns="http://www.idpf.org/2007/opf" property="file-as" refines="#author-name">Writer, Example</meta>
            """));
        byte[] input = EpubIntegrityFixtures.ReplaceEntry(book.Write().Bytes, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()));
        var creator = Assert.Single(EpubDocument.Load(new MemoryStream(input)).Metadata, entry => entry.Id == "author-name");
        Assert.Equal("aut", creator.Role);
        Assert.Equal("Writer, Example", creator.FileAs);
        metadata.Elements().Single(element => (string?)element.Attribute("id") == "author-name").SetAttributeValue(opf + "role", "edt");
        input = EpubIntegrityFixtures.ReplaceEntry(input, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()));
        creator = Assert.Single(EpubDocument.Load(new MemoryStream(input)).Metadata, entry => entry.Id == "author-name");
        Assert.Equal("edt", creator.Role);
    }

    [Fact]
    public void ContributorRecordsRetainNamesRolesLanguagesAndExistingMetadata() {
        var book = Book();
        book.Creator = "Existing author";
        book.DeclareVocabularyPrefix("custom", "https://example.org/metadata/");
        book.AddMetadataProperty("custom:editorial-status", "reviewed");
        book.AddCreator("alice", new EpubContributorMetadata {
            Name = "Alice Example", FileAs = "Example, Alice", Language = "en",
            MarcRoles = new[] { "aut", "ill", "aut" }
        });
        book.AddContributor("translator", new EpubContributorMetadata {
            Name = "Jan Kowalski", Language = "pl", MarcRoles = new[] { "trl" }
        });
        EpubDocument reopened = EpubDocument.Load(new MemoryStream(book.Write().Bytes));
        var creator = Assert.Single(reopened.Metadata, entry => entry.Id == "alice");
        Assert.Equal("creator", creator.Name);
        Assert.Equal("Alice Example", creator.Value);
        Assert.Equal("Example, Alice", creator.FileAs);
        Assert.Equal("en", creator.Language);
        var roles = reopened.Metadata.Where(entry => entry.Refines == "#alice" && entry.Property == "role").ToArray();
        Assert.Equal(new[] { "aut", "ill" }, roles.Select(entry => entry.Value));
        Assert.All(roles, entry => Assert.Equal("marc:relators", entry.Scheme));
        var translator = Assert.Single(reopened.Metadata, entry => entry.Id == "translator");
        Assert.Equal("contributor", translator.Name);
        Assert.Equal("trl", translator.Role);
        Assert.Equal("pl", translator.Language);
        Assert.Contains(reopened.Metadata, entry => entry.Name == "creator" && entry.Value == "Existing author");
        Assert.Contains(reopened.Metadata, entry => entry.Property == "custom:editorial-status" && entry.Value == "reviewed");
        int creatorIndex = reopened.Metadata.ToList().FindIndex(entry => entry.Id == "alice");
        var limited = EpubDocument.Load(new MemoryStream(book.Write().Bytes), new EpubReadOptions { MaxMetadataItems = creatorIndex + 1 });
        var limitedCreator = Assert.Single(limited.Metadata, entry => entry.Id == "alice");
        Assert.Null(limitedCreator.Role);
        Assert.Null(limitedCreator.FileAs);
        Assert.Contains(limited.Diagnostics, entry => entry.Code == "epub.metadata.count-limit");
    }

    [Fact]
    public void CollectionsPreserveHierarchicalPositionsAndIndependentMemberships() {
        var book = Book();
        book.AddCollection("series", new EpubCollectionMetadata {
            Name = "The Example Chronicles", Kind = EpubCollectionKind.Series,
            FileAs = "Example Chronicles, The", Language = "en", Position = new uint[] { 2, 10, 1 }
        });
        book.AddCollection("set", new EpubCollectionMetadata { Name = "Collected Works", Kind = EpubCollectionKind.Set });
        var metadata = EpubDocument.Load(new MemoryStream(book.Write().Bytes)).Metadata;
        var series = Assert.Single(metadata, item => item.Id == "series");
        Assert.Equal("belongs-to-collection", series.Property);
        Assert.Equal("en", series.Language);
        Assert.Contains(metadata, item => item.Refines == "#series" && item.Property == "group-position" && item.Value == "2.10.1");
        Assert.Contains(metadata, item => item.Refines == "#series" && item.Property == "collection-type" && item.Value == "series");
        Assert.Contains(metadata, item => item.Refines == "#series" && item.Property == "file-as" && item.Value == "Example Chronicles, The");
        Assert.Contains(metadata, item => item.Refines == "#set" && item.Property == "collection-type" && item.Value == "set");
        Assert.DoesNotContain(metadata, item => item.Refines == "#set" && item.Property == "group-position");
    }

    [Fact]
    public void InvalidRefinedRecordsLeaveNoPartialMetadata() {
        var book = Book();
        byte[] original = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.AddCreator("chapter", new EpubContributorMetadata { Name = "Duplicate package id" }));
        Assert.Throws<ArgumentException>(() => book.AddContributor("bad-role", new EpubContributorMetadata { Name = "Name", MarcRoles = new[] { "aut", "Author" } }));
        Assert.Throws<ArgumentException>(() => book.AddContributor("bad-lang", new EpubContributorMetadata { Name = "Name", Language = "English language" }));
        Assert.Throws<ArgumentOutOfRangeException>(() => book.AddCollection("bad-kind", new EpubCollectionMetadata { Name = "Name", Kind = (EpubCollectionKind)99 }));
        Assert.Equal(original, book.Write().Bytes);
        book.DeclareVocabularyPrefix("marc", "https://example.org/not-marc/");
        byte[] remapped = book.Write().Bytes;
        Assert.Throws<InvalidOperationException>(() => book.AddCreator("author", new EpubContributorMetadata { Name = "Name", MarcRoles = new[] { "aut" } }));
        Assert.Equal(remapped, book.Write().Bytes);
    }

    [Fact]
    public void CombinedRecordBudgetFailureIsAtomicAndLegacyVersionIsExplicit() {
        byte[] input = Book().Write().Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var bounded = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxExpandedBytes = zip.Entries.Sum(entry => entry.Length) + 100 });
        Assert.Throws<InvalidDataException>(() => bounded.AddCreator("new-author", new EpubContributorMetadata {
            Name = "Name", MarcRoles = new[] { "aut", "ill" }, FileAs = "Sorting name"
        }));
        Assert.Equal(input, bounded.Write().Bytes);
        Assert.Throws<InvalidDataException>(() => bounded.AddCollection("new-series", new EpubCollectionMetadata {
            Name = "Name", Position = new uint[] { 1 }, FileAs = "Sorting name"
        }));
        Assert.Equal(input, bounded.Write().Bytes);
        var legacy = EpubPublication.Create("Legacy", "en", version: EpubVersion.Epub2);
        Assert.Throws<NotSupportedException>(() => legacy.AddCreator("author", new EpubContributorMetadata { Name = "Name" }));
        Assert.Throws<NotSupportedException>(() => legacy.AddCollection("series", new EpubCollectionMetadata { Name = "Name" }));
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Publishing metadata", "en");
        book.AddChapter("chapter", "EPUB/chapter.xhtml", "Chapter", "<h1>Chapter</h1><p>Text.</p>");
        return book;
    }
}

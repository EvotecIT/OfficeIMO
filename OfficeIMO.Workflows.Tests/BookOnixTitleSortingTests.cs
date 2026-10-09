using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixTitleSortingTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;

    [Theory]
    [InlineData("The history & future", "The ", "history & future")]
    [InlineData("L’histoire", "L’", "histoire")]
    [InlineData("The\u00a0history", "The\u00a0", "history")]
    [InlineData("The  history", "The ", " history")]
    public void ProductPrefixUsesExactSelectedExportedTitleAndPreservesItsCharacters(string title, string prefix, string remainder) {
        var project = BookOnixTests.Project();
        project.Publication.AddTitle("selected", new() { Text = title });
        var before = project.ToProjectBytes();
        var options = BookOnixTests.Options() with { TitleId = "selected", Subtitle = "A subtitle" };
        var baseline = project.ExportOnix(options, BookOnixTests.TestSchema());
        var result = project.ExportOnix(options with { TitleSorting = new() { Prefix = prefix } }, BookOnixTests.TestSchema());
        var element = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "TitleElement").Single();
        Assert.Equal(new[] { "TitleElementLevel", "TitlePrefix", "TitleWithoutPrefix", "Subtitle" }, element.Elements().Select(e => e.Name.LocalName));
        Assert.Equal(prefix, element.Element(Ns + "TitlePrefix")!.Value);
        Assert.Equal(remainder, element.Element(Ns + "TitleWithoutPrefix")!.Value);
        Assert.Equal(title, element.Element(Ns + "TitlePrefix")!.Value + element.Element(Ns + "TitleWithoutPrefix")!.Value);
        Assert.Equal("A subtitle", element.Element(Ns + "Subtitle")!.Value);
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void OmittedSortingAndExplicitNoPrefixProduceDistinctDeclarations() {
        var project = BookOnixTests.Project();
        var omitted = XDocument.Load(new MemoryStream(project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema()).Bytes));
        Assert.Single(omitted.Descendants(Ns + "TitleText"));
        Assert.Empty(omitted.Descendants(Ns + "NoPrefix"));
        var explicitNone = XDocument.Load(new MemoryStream(project.ExportOnix(BookOnixTests.Options() with {
            TitleSorting = new() }, BookOnixTests.TestSchema()).Bytes));
        Assert.Empty(explicitNone.Descendants(Ns + "TitleText"));
        Assert.Single(explicitNone.Descendants(Ns + "NoPrefix"));
        Assert.Equal("Book & title", explicitNone.Descendants(Ns + "TitleWithoutPrefix").Single().Value);
    }

    [Fact]
    public void SimpleAndHierarchicalCollectionTitlesUseTheirOwnSortingAndLanguage() {
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Collections = [
            new() { Type = BookOnixCollectionType.Publisher, Title = "Les études", LanguageCode = "fre", TitleSorting = new() { Prefix = "Les " } },
            new() { Type = BookOnixCollectionType.Publisher, TitleElements = [
                new() { Level = BookOnixCollectionLevel.Collection, Title = "The studies", LanguageCode = "eng", TitleSorting = new() },
                new() { Level = BookOnixCollectionLevel.Subcollection, Title = "L’histoire", LanguageCode = "fre", PartNumber = "II", TitleSorting = new() { Prefix = "L’" } }
            ] }
        ] }, BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        var titles = xml.Descendants(Ns + "Collection").SelectMany(c => c.Descendants(Ns + "TitleElement")).ToArray();
        Assert.Equal("Les ", titles[0].Element(Ns + "TitlePrefix")!.Value);
        Assert.Equal("études", titles[0].Element(Ns + "TitleWithoutPrefix")!.Value);
        Assert.Equal("fre", (string?)titles[0].Element(Ns + "TitlePrefix")!.Attribute("language"));
        Assert.Equal("fre", (string?)titles[0].Element(Ns + "TitleWithoutPrefix")!.Attribute("language"));
        Assert.NotNull(titles[1].Element(Ns + "NoPrefix"));
        Assert.Equal("The studies", titles[1].Element(Ns + "TitleWithoutPrefix")!.Value);
        Assert.Equal("eng", (string?)titles[1].Element(Ns + "TitleWithoutPrefix")!.Attribute("language"));
        Assert.Equal(new[] { "SequenceNumber", "TitleElementLevel", "PartNumber", "TitlePrefix", "TitleWithoutPrefix" }, titles[2].Elements().Select(e => e.Name.LocalName));
        Assert.Equal("L’", titles[2].Element(Ns + "TitlePrefix")!.Value);
        Assert.Equal("histoire", titles[2].Element(Ns + "TitleWithoutPrefix")!.Value);
        Assert.Equal("Book & title", xml.Descendants(Ns + "DescriptiveDetail").Single().Elements(Ns + "TitleDetail").Single().Descendants(Ns + "TitleText").Single().Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("book")]
    [InlineData("Book & title")]
    [InlineData("Book & title longer")]
    public void InvalidOrExhaustivePrefixFailsWithoutChangingProject(string prefix) {
        var project = BookOnixTests.Project();
        var before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with {
            TitleSorting = new() { Prefix = prefix } }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void SortingCannotBeAttachedToPartOnlyTitlesOrDiscardedByTheHierarchy() {
        var partOnly = new BookOnixCollection { Type = BookOnixCollectionType.Publisher,
            TitleElements = [new() { Level = BookOnixCollectionLevel.Collection, PartNumber = "Part 1", TitleSorting = new() }] };
        var ambiguous = new BookOnixCollection { Type = BookOnixCollectionType.Publisher, TitleSorting = new(),
            TitleElements = [new() { Level = BookOnixCollectionLevel.Collection, Title = "Series" }] };
        foreach (var collection in new[] { partOnly, ambiguous })
            Assert.ThrowsAny<ArgumentException>(() => BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with {
                Collections = [collection] }, BookOnixTests.TestSchema()));
    }
}

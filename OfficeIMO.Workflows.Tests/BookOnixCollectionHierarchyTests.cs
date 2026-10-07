using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixCollectionHierarchyTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixCollectionTitleElement Title(BookOnixCollectionLevel level) => new() { Level = level, Title = level.ToString() };
    private static BookOnixCollection Hierarchy() => new() {
        Type = BookOnixCollectionType.Ascribed, SourceName = "Library", Frequency = BookOnixCollectionFrequency.Monthly,
        TitleElements = [Title(BookOnixCollectionLevel.Subcollection) with { Title = "Études & histoire", Subtitle = "Archives",
            PartNumber = "Volume II", LanguageCode = "fre" },
            Title(BookOnixCollectionLevel.Collection) with { Title = "Collected studies", LanguageCode = "eng" },
            Title(BookOnixCollectionLevel.SubSubcollection) with { Title = null, PartNumber = "Part 3" }],
        Identifiers = [new(BookOnixCollectionIdentifierType.Issn, "0317-8471") { Level = BookOnixCollectionLevel.Collection },
            new(BookOnixCollectionIdentifierType.Issn, "1092-003X") { Level = BookOnixCollectionLevel.Subcollection }]
    };

    [Fact]
    public void HierarchyKeepsDisplayOrderLanguageAndIdentifierScopeWithoutChangingPublication() {
        var project = BookOnixTests.Project();
        var before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { Collections = [Hierarchy()] }, BookOnixTests.TestSchema());
        var document = XDocument.Load(new MemoryStream(result.Bytes));
        var collection = document.Descendants(Ns + "Collection").Single();
        Assert.Equal(new[] { "CollectionType", "CollectionFrequency", "SourceName", "CollectionIdentifier", "CollectionIdentifier", "TitleDetail" },
            collection.Elements().Select(e => e.Name.LocalName));
        Assert.Equal("m", collection.Element(Ns + "CollectionFrequency")!.Value);
        var titles = collection.Descendants(Ns + "TitleElement").ToArray();
        Assert.Equal(new[] { "1", "2", "3" }, titles.Select(e => e.Element(Ns + "SequenceNumber")!.Value));
        Assert.Equal(new[] { "03", "02", "06" }, titles.Select(e => e.Element(Ns + "TitleElementLevel")!.Value));
        Assert.Equal(new[] { "SequenceNumber", "TitleElementLevel", "PartNumber", "TitleText", "Subtitle" }, titles[0].Elements().Select(e => e.Name.LocalName));
        Assert.Equal("Études & histoire", titles[0].Element(Ns + "TitleText")!.Value);
        Assert.All(titles[0].Elements().Skip(2), e => Assert.Equal("fre", (string?)e.Attribute("language")));
        Assert.Equal("eng", (string?)titles[1].Element(Ns + "TitleText")!.Attribute("language"));
        Assert.Null(titles[2].Element(Ns + "TitleText"));
        Assert.Null(titles[2].Element(Ns + "PartNumber")!.Attribute("language"));
        Assert.Equal("Part 3", titles[2].Element(Ns + "PartNumber")!.Value);
        var ids = collection.Elements(Ns + "CollectionIdentifier").ToArray();
        Assert.Equal(new[] { "02", "03" }, ids.Select(e => e.Elements().First().Value));
        Assert.Equal(new[] { "03178471", "1092003X" }, ids.Select(e => e.Element(Ns + "IDValue")!.Value));
        Assert.Equal("Book & title", document.Descendants(Ns + "DescriptiveDetail").Single().Elements(Ns + "TitleDetail").Single().Descendants(Ns + "TitleText").Single().Value);
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void FrequencyOmissionUnknownAndEndOfCollectionRemainDistinct() {
        var project = BookOnixTests.Project();
        var options = BookOnixTests.Options() with { Collections = [new() { Type = BookOnixCollectionType.Publisher, Title = "Series" },
            new() { Type = BookOnixCollectionType.Publisher, Title = "Unknown", Frequency = BookOnixCollectionFrequency.Unknown },
            new() { Type = BookOnixCollectionType.Publisher, Title = "Ended", Frequency = BookOnixCollectionFrequency.NoFuturePublications }] };
        var collections = XDocument.Load(new MemoryStream(project.ExportOnix(options, BookOnixTests.TestSchema()).Bytes)).Descendants(Ns + "Collection").ToArray();
        Assert.Null(collections[0].Element(Ns + "CollectionFrequency"));
        Assert.Null(collections[0].Descendants(Ns + "TitleElement").Single().Element(Ns + "SequenceNumber"));
        Assert.Equal("u", collections[1].Element(Ns + "CollectionFrequency")!.Value);
        Assert.Equal("x", collections[2].Element(Ns + "CollectionFrequency")!.Value);
    }

    [Theory]
    [InlineData("null-elements")]
    [InlineData("null-element")]
    [InlineData("title-conflict")]
    [InlineData("subtitle-conflict")]
    [InlineData("language-conflict")]
    [InlineData("too-many-elements")]
    [InlineData("duplicate-level")]
    [InlineData("missing-root")]
    [InlineData("missing-parent")]
    [InlineData("invalid-level")]
    [InlineData("missing-title-or-part")]
    [InlineData("empty-title")]
    [InlineData("empty-part")]
    [InlineData("long-part")]
    [InlineData("invalid-language")]
    [InlineData("identifier-missing-level")]
    [InlineData("identifier-invalid-level")]
    [InlineData("duplicate-identifier")]
    [InlineData("unscoped-overlap-first")]
    [InlineData("unscoped-overlap-last")]
    [InlineData("invalid-frequency")]
    public void InvalidHierarchyAssertionsFailWithoutMutation(string kind) {
        var root = Title(BookOnixCollectionLevel.Collection);
        var child = Title(BookOnixCollectionLevel.Subcollection);
        var identifier = new BookOnixCollectionIdentifier(BookOnixCollectionIdentifierType.Issn, "03178471");
        var collection = kind switch {
            "null-elements" => Hierarchy() with { TitleElements = null! },
            "null-element" => Hierarchy() with { TitleElements = [null!] },
            "title-conflict" => Hierarchy() with { Title = "Ambiguous" },
            "subtitle-conflict" => Hierarchy() with { Subtitle = "Ambiguous" },
            "language-conflict" => Hierarchy() with { LanguageCode = "eng" },
            "too-many-elements" => Hierarchy() with { TitleElements = [root, child, root, child, root, child] },
            "duplicate-level" => Hierarchy() with { TitleElements = [root, root] },
            "missing-root" => Hierarchy() with { TitleElements = [child] },
            "missing-parent" => Hierarchy() with { TitleElements = [root, Title(BookOnixCollectionLevel.SubSubcollection)] },
            "invalid-level" => Hierarchy() with { TitleElements = [root, Title((BookOnixCollectionLevel)99)] },
            "missing-title-or-part" => Hierarchy() with { TitleElements = [root, child with { Title = null }] },
            "empty-title" => Hierarchy() with { TitleElements = [root, child with { Title = " " }] },
            "empty-part" => Hierarchy() with { TitleElements = [root, child with { PartNumber = " " }] },
            "long-part" => Hierarchy() with { TitleElements = [root, child with { PartNumber = new string('a', 4097) }] },
            "invalid-language" => Hierarchy() with { TitleElements = [root, child with { LanguageCode = "en-US" }] },
            "identifier-missing-level" => Hierarchy() with { TitleElements = [root] },
            "identifier-invalid-level" => Hierarchy() with { Identifiers = [identifier with { Level = (BookOnixCollectionLevel)99 }] },
            "duplicate-identifier" => Hierarchy() with { Identifiers = [identifier with { Level = BookOnixCollectionLevel.Collection }, identifier with { Level = BookOnixCollectionLevel.Collection }] },
            "unscoped-overlap-first" => Hierarchy() with { Identifiers = [identifier, identifier with { Level = BookOnixCollectionLevel.Collection }] },
            "unscoped-overlap-last" => Hierarchy() with { Identifiers = [identifier with { Level = BookOnixCollectionLevel.Collection }, identifier] },
            _ => Hierarchy() with { Frequency = (BookOnixCollectionFrequency)99 }
        };
        var project = BookOnixTests.Project();
        var before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Collections = [collection] }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

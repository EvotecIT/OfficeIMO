using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixResourceContributorTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixContributorIdentifier Identity(BookOnixContributorIdentifierType type) => type switch {
        BookOnixContributorIdentifierType.Isni => new(type, "0000000121032683"),
        BookOnixContributorIdentifierType.Orcid => new(type, "0000000218250097"),
        _ => new(type, "author-1", "Publisher people")
    };
    private static BookOnixExportOptions Options(BookOnixContributorIdentifier identity) => BookOnixTests.Options() with {
        Contributors = [new("Writer", BookOnixContributorRole.Author) { Identifiers = [identity] }],
        SupportingResources = [new() {
            Type = BookOnixResourceContentType.FrontCover, Mode = BookOnixResourceMode.Image,
            Audiences = [BookOnixContentAudience.Unrestricted], ContributorReferences = [identity],
            Versions = [new() { Form = BookOnixResourceForm.Linkable, Links = [new("https://example.org/portrait.jpg")] }]
        }]
    };

    [Theory]
    [InlineData(BookOnixContributorIdentifierType.Isni, "16", "05")]
    [InlineData(BookOnixContributorIdentifierType.Orcid, "21", "11")]
    [InlineData(BookOnixContributorIdentifierType.Proprietary, "01", "06")]
    public void ResourceIdentityMatchesExportedContributorAndPreservesPublication(BookOnixContributorIdentifierType type, string nameType, string featureType) {
        var project = BookOnixTests.Project(); var identity = Identity(type);
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(Options(identity), BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        var contributor = xml.Descendants(Ns + "Contributor").Single();
        Assert.Equal(new[] { "SequenceNumber", "ContributorRole", "NameIdentifier", "PersonName" }, contributor.Elements().Select(x => x.Name.LocalName));
        Assert.Equal(nameType, contributor.Descendants(Ns + "NameIDType").Single().Value);
        Assert.Equal(identity.SchemeName, (string?)contributor.Descendants(Ns + "IDTypeName").SingleOrDefault());
        var feature = xml.Descendants(Ns + "ResourceFeature").Single();
        Assert.Equal(featureType, feature.Element(Ns + "ResourceFeatureType")!.Value);
        Assert.Equal(identity.Value, feature.Element(Ns + "FeatureValue")!.Value);
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
    }

    [Fact]
    public void OneIdentityMayAppearInSeveralContributorRoles() {
        var identity = Identity(BookOnixContributorIdentifierType.Isni);
        var options = Options(identity);
        options = options with { Contributors = [options.Contributors[0], options.Contributors[0] with { Role = BookOnixContributorRole.Illustrator }] };
        var xml = XDocument.Load(new MemoryStream(BookOnixTests.Project().ExportOnix(options, BookOnixTests.TestSchema()).Bytes));
        Assert.Equal(2, xml.Descendants(Ns + "Contributor").Count());
        Assert.Single(xml.Descendants(Ns + "ResourceFeature"));
    }

    [Fact]
    public void CollectionIdentityDoesNotSatisfyProductResourceReference() {
        var identity = Identity(BookOnixContributorIdentifierType.Isni);
        var options = Options(identity);
        options = options with { Contributors = [new("Other writer", BookOnixContributorRole.Author)],
            Collections = [new() { Type = BookOnixCollectionType.Publisher, Title = "Series", Contributors = [options.Contributors[0]] }] };
        Assert.Throws<ArgumentException>(() => BookOnixTests.Project().ExportOnix(options, BookOnixTests.TestSchema()));
        var withoutResource = options with { SupportingResources = [] };
        var xml = XDocument.Load(new MemoryStream(BookOnixTests.Project().ExportOnix(withoutResource, BookOnixTests.TestSchema()).Bytes));
        Assert.Equal(identity.Value, xml.Descendants(Ns + "Collection").Single().Descendants(Ns + "IDValue").Single().Value);
    }

    [Theory]
    [InlineData("missing")][InlineData("scheme-mismatch")][InlineData("scheme-collision")][InlineData("duplicate-reference")]
    [InlineData("null-references")][InlineData("null-reference")][InlineData("too-many-references")]
    [InlineData("null-identifiers")][InlineData("null-identifier")][InlineData("too-many-identifiers")]
    [InlineData("duplicate-scheme")][InlineData("missing-scheme")][InlineData("long-scheme")][InlineData("long-value")]
    [InlineData("xml-value")][InlineData("empty-value")][InlineData("standard-scheme")][InlineData("separators")]
    [InlineData("nonascii")][InlineData("lowercase-check")][InlineData("unknown-type")]
    public void InvalidIdentityOrRelationshipFailsWithoutMutation(string kind) {
        var identity = Identity(BookOnixContributorIdentifierType.Proprietary);
        var options = Options(identity);
        var invalid = kind switch {
            "missing-scheme" => identity with { SchemeName = null }, "long-scheme" => identity with { SchemeName = new string('a', 101) },
            "long-value" => identity with { Value = new string('a', 101) }, "xml-value" => identity with { Value = "bad\u0001" },
            "empty-value" => identity with { Value = " " }, "standard-scheme" => identity with { Type = BookOnixContributorIdentifierType.Isni },
            "separators" => Identity(BookOnixContributorIdentifierType.Orcid) with { Value = "0000-0002-1825-0097" },
            "nonascii" => Identity(BookOnixContributorIdentifierType.Isni) with { Value = new string('١', 16) },
            "lowercase-check" => Identity(BookOnixContributorIdentifierType.Isni) with { Value = "000000012103268x" },
            "unknown-type" => identity with { Type = (BookOnixContributorIdentifierType)99 }, _ => identity
        };
        var identifiers = kind switch {
            "missing" => Array.Empty<BookOnixContributorIdentifier>(), "null-identifiers" => null!,
            "null-identifier" => new BookOnixContributorIdentifier[] { null! },
            "too-many-identifiers" => Enumerable.Repeat(identity, 17).ToArray(),
            "duplicate-scheme" => new[] { identity, identity with { Value = "other" } },
            "scheme-mismatch" => new[] { identity with { SchemeName = "Other" } },
            "scheme-collision" => new[] { identity, identity with { SchemeName = "Other" } }, _ => new[] { invalid }
        };
        var references = kind switch {
            "duplicate-reference" => new[] { identity, identity }, "null-references" => null!,
            "null-reference" => new BookOnixContributorIdentifier[] { null! },
            "too-many-references" => Enumerable.Repeat(identity, 17).ToArray(), _ => new[] { identity }
        };
        options = options with { Contributors = [options.Contributors[0] with { Identifiers = identifiers }],
            SupportingResources = [options.SupportingResources[0] with { ContributorReferences = references }] };
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        if (kind == "xml-value") Assert.Throws<System.Xml.XmlException>(() => project.ExportOnix(options, BookOnixTests.TestSchema()));
        else Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(options, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

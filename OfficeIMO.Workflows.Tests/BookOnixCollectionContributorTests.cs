using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixCollectionContributorTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixCollection Collection() => new() { Type = BookOnixCollectionType.Publisher, Title = "Collected studies" };

    [Fact]
    public void CollectionCreditsKeepTheirOwnOrderAndDoNotChangeProductOrEpubCredits() {
        var project = BookOnixTests.Project();
        var before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { Collections = [Collection() with {
            Contributors = [new("Editor & colleague", BookOnixContributorRole.SeriesEditor),
                new("Editorial Studio", BookOnixContributorRole.Editor, true)] },
            Collection() with { Title = "Other series", Contributors = [new("Another editor", BookOnixContributorRole.SeriesEditor)] }]
        }, BookOnixTests.TestSchema());
        var descriptive = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "DescriptiveDetail").Single();
        var collections = descriptive.Elements(Ns + "Collection").ToArray();
        var credits = collections[0].Elements(Ns + "Contributor").ToArray();
        Assert.Equal(new[] { "CollectionType", "TitleDetail", "Contributor", "Contributor" }, collections[0].Elements().Select(e => e.Name.LocalName));
        Assert.Equal(new[] { "1", "2" }, credits.Select(e => e.Element(Ns + "SequenceNumber")!.Value));
        Assert.Equal(new[] { "B09", "B01" }, credits.Select(e => e.Element(Ns + "ContributorRole")!.Value));
        Assert.Equal("Editor & colleague", credits[0].Element(Ns + "PersonName")!.Value);
        Assert.Equal("Editorial Studio", credits[1].Element(Ns + "CorporateName")!.Value);
        Assert.Null(credits[1].Element(Ns + "PersonName"));
        Assert.Equal("1", collections[1].Element(Ns + "Contributor")!.Element(Ns + "SequenceNumber")!.Value);
        Assert.Equal("Author", descriptive.Elements(Ns + "Contributor").Single().Element(Ns + "PersonName")!.Value);
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void OmittedAndExplicitlyAbsentCollectionCreditsAreDistinct() {
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Collections = [Collection(),
            Collection() with { NoContributors = true }] }, BookOnixTests.TestSchema());
        var collections = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "Collection").ToArray();
        Assert.Empty(collections[0].Elements(Ns + "Contributor"));
        Assert.Null(collections[0].Element(Ns + "NoContributor"));
        Assert.Equal("NoContributor", collections[1].Elements().Last().Name.LocalName);
        Assert.Empty(collections[1].Elements(Ns + "Contributor"));
    }

    [Fact]
    public void ProductCreditsAlsoSupportExplicitSeriesEditorRole() {
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with {
            Contributors = [new("Series editor", BookOnixContributorRole.SeriesEditor)] }, BookOnixTests.TestSchema());
        Assert.Equal("B09", XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "ContributorRole").Single().Value);
    }

    [Theory]
    [InlineData("null-list")]
    [InlineData("null-credit")]
    [InlineData("blank-name")]
    [InlineData("long-name")]
    [InlineData("invalid-role")]
    [InlineData("too-many")]
    [InlineData("conflicting-assertion")]
    public void InvalidCollectionCreditsFailWithoutMutation(string kind) {
        var credit = new BookOnixContributor("Editor", BookOnixContributorRole.SeriesEditor);
        var collection = kind switch {
            "null-list" => Collection() with { Contributors = null! },
            "null-credit" => Collection() with { Contributors = [null!] },
            "blank-name" => Collection() with { Contributors = [credit with { Name = " " }] },
            "long-name" => Collection() with { Contributors = [credit with { Name = new string('a', 4097) }] },
            "invalid-role" => Collection() with { Contributors = [credit with { Role = (BookOnixContributorRole)99 }] },
            "too-many" => Collection() with { Contributors = Enumerable.Repeat(credit, 101).ToArray() },
            _ => Collection() with { Contributors = [credit], NoContributors = true }
        };
        var project = BookOnixTests.Project();
        var before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with {
            Collections = [collection] }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

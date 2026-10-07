using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixEditionTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixEdition Edition() => new() { Number = 2, VersionNumber = "1.2", Types = [BookOnixEditionType.Revised, BookOnixEditionType.Annotated],
        Statements = [new("Second revised & annotated edition", "eng"), new("Drugie wydanie", "pol")] };

    [Fact]
    public void EditionDetailsRemainExplicitOrderedAndBoundToAnUnchangedPublication() {
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { Edition = Edition() }, BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        var detail = xml.Descendants(Ns + "DescriptiveDetail").Single();
        Assert.Equal(new[] { "REV", "ANN" }, detail.Elements(Ns + "EditionType").Select(e => e.Value));
        Assert.Equal("2", detail.Element(Ns + "EditionNumber")!.Value);
        Assert.Equal("1.2", detail.Element(Ns + "EditionVersionNumber")!.Value);
        var statements = detail.Elements(Ns + "EditionStatement").ToArray();
        Assert.Equal(new[] { "eng", "pol" }, statements.Select(e => (string?)e.Attribute("language")));
        Assert.All(statements, e => Assert.Equal("06", (string?)e.Attribute("textformat")));
        Assert.Equal("Second revised & annotated edition", statements[0].Value);
        Assert.Equal("Language", statements.Last().ElementsAfterSelf().First().Name.LocalName);
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void NoEditionIsDistinctFromOmittedEditionMetadata() {
        var project = BookOnixTests.Project();
        var baseline = XDocument.Load(new MemoryStream(project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema()).Bytes));
        var explicitNone = XDocument.Load(new MemoryStream(project.ExportOnix(BookOnixTests.Options() with { Edition = new() { NoEdition = true } }, BookOnixTests.TestSchema()).Bytes));
        Assert.Empty(baseline.Descendants(Ns + "NoEdition"));
        Assert.Single(explicitNone.Descendants(Ns + "NoEdition"));
    }

    [Theory]
    [InlineData("empty")]
    [InlineData("contradiction")]
    [InlineData("zero")]
    [InlineData("negative")]
    [InlineData("orphan-version")]
    [InlineData("empty-version")]
    [InlineData("unknown-type")]
    [InlineData("duplicate-type")]
    [InlineData("contradictory-types")]
    [InlineData("new-numbered")]
    [InlineData("new-specific")]
    [InlineData("language")]
    [InlineData("duplicate-language")]
    [InlineData("empty-statement")]
    [InlineData("statement-count")]
    public void InvalidEditionAssertionsFailWithoutProjectMutation(string kind) {
        var edition = kind switch {
            "empty" => new BookOnixEdition(),
            "contradiction" => Edition() with { NoEdition = true },
            "zero" => Edition() with { Number = 0 },
            "negative" => Edition() with { Number = -1 },
            "orphan-version" => Edition() with { Number = null },
            "empty-version" => Edition() with { VersionNumber = " " },
            "unknown-type" => Edition() with { Types = [(BookOnixEditionType)99] },
            "duplicate-type" => Edition() with { Types = [BookOnixEditionType.Revised, BookOnixEditionType.Revised] },
            "contradictory-types" => Edition() with { Types = [BookOnixEditionType.Abridged, BookOnixEditionType.Unabridged] },
            "new-numbered" => Edition() with { Types = [BookOnixEditionType.New] },
            "new-specific" => new() { Types = [BookOnixEditionType.New, BookOnixEditionType.Revised] },
            "language" => Edition() with { Statements = [new("Edition", "en")] },
            "duplicate-language" => Edition() with { Statements = [new("One"), new("Two")] },
            "empty-statement" => Edition() with { Statements = [new(" ")] },
            _ => Edition() with { Statements = Enumerable.Repeat(new BookOnixEditionStatement("Edition"), 17).ToArray() }
        };
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Edition = edition }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

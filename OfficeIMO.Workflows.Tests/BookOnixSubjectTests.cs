using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixSubjectTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixSubject Subject() => new() { Scheme = BookOnixSubjectScheme.Thema, Code = "YFB", IsMain = true,
        SchemeVersion = "1.6", Headings = [new("Children's fiction", "eng"), new("Literatura dziecięca", "pol")] };

    [Fact]
    public void SelectedTitleAndMultilingualSubjectsPreserveExplicitMetadata() {
        var project = BookOnixTests.Project();
        project.Publication.AddTitle("alternate", new() { Text = "Alternate & selected title", Kind = EpubTitleKind.Main });
        byte[] before = project.ToProjectBytes();
        var result = project.ExportOnix(BookOnixTests.Options() with { TitleId = "alternate", Subjects = [Subject(),
            new() { Scheme = BookOnixSubjectScheme.Bisac, Code = "JUV000000", IsMain = true },
            new() { Scheme = BookOnixSubjectScheme.Keywords, Headings = [new("stories; adventure")] },
            new() { Scheme = BookOnixSubjectScheme.Proprietary, SchemeName = "Example & partners", Code = "A1", IsMain = true }]
        }, BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        Assert.Equal("Alternate & selected title", xml.Descendants(Ns + "TitleText").Single().Value);
        var subjects = xml.Descendants(Ns + "Subject").ToArray();
        Assert.Equal(new[] { "93", "10", "20", "24" }, subjects.Select(s => s.Element(Ns + "SubjectSchemeIdentifier")!.Value));
        Assert.Equal("1.6", subjects[0].Element(Ns + "SubjectSchemeVersion")!.Value);
        Assert.Equal(new[] { "eng", "pol" }, subjects[0].Elements(Ns + "SubjectHeadingText").Select(e => (string?)e.Attribute("language")));
        Assert.Equal("Literatura dziecięca", subjects[0].Elements(Ns + "SubjectHeadingText").Last().Value);
        Assert.Equal(3, xml.Descendants(Ns + "MainSubject").Count());
        Assert.Equal(before, project.ToProjectBytes());
        var original = XDocument.Load(new MemoryStream(project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema()).Bytes));
        Assert.Equal("Book & title", original.Descendants(Ns + "TitleText").Single().Value);
        Assert.Throws<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { TitleId = "missing" }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Theory]
    [InlineData("empty")]
    [InlineData("scheme")]
    [InlineData("name")]
    [InlineData("standard-name")]
    [InlineData("language")]
    [InlineData("duplicate-language")]
    [InlineData("duplicate-unspecified")]
    [InlineData("duplicate-main")]
    [InlineData("keyword-code")]
    [InlineData("keyword-main")]
    [InlineData("qualifier-main")]
    [InlineData("heading-count")]
    [InlineData("subject-count")]
    public void InvalidClassificationsAreRejectedWithoutMutation(string kind) {
        var subject = kind switch {
            "empty" => Subject() with { Code = null, Headings = [] },
            "scheme" => Subject() with { Scheme = (BookOnixSubjectScheme)99 },
            "name" => Subject() with { Scheme = BookOnixSubjectScheme.Proprietary },
            "standard-name" => Subject() with { SchemeName = "Custom" },
            "language" => Subject() with { Headings = [new("Text", "en-US")] },
            "duplicate-language" => Subject() with { Headings = [new("A", "eng"), new("B", "eng")] },
            "duplicate-unspecified" => Subject() with { Headings = [new("A"), new("B")] },
            "keyword-code" => Subject() with { Scheme = BookOnixSubjectScheme.Keywords, IsMain = false },
            "keyword-main" => Subject() with { Scheme = BookOnixSubjectScheme.Keywords, Code = null },
            "qualifier-main" => Subject() with { Scheme = BookOnixSubjectScheme.ThemaGeographical },
            "heading-count" => Subject() with { Headings = Enumerable.Repeat(new BookOnixSubjectHeading("Text"), 17).ToArray() },
            _ => Subject()
        };
        var subjects = kind == "duplicate-main" ? new[] { subject, subject with { Code = "YFC" } } :
            kind == "subject-count" ? Enumerable.Repeat(subject, 65).ToArray() : [subject];
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Subjects = subjects }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

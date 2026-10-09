using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixAlternativeTitleTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;

    [Theory]
    [InlineData(BookOnixAlternativeTitleType.OriginalLanguage, "03")]
    [InlineData(BookOnixAlternativeTitleType.Abbreviated, "05")]
    [InlineData(BookOnixAlternativeTitleType.OtherLanguage, "06")]
    [InlineData(BookOnixAlternativeTitleType.Former, "08")]
    [InlineData(BookOnixAlternativeTitleType.Distributor, "10")]
    [InlineData(BookOnixAlternativeTitleType.Cover, "11")]
    [InlineData(BookOnixAlternativeTitleType.BackCover, "12")]
    [InlineData(BookOnixAlternativeTitleType.Expanded, "13")]
    [InlineData(BookOnixAlternativeTitleType.Alternative, "14")]
    [InlineData(BookOnixAlternativeTitleType.Spine, "15")]
    [InlineData(BookOnixAlternativeTitleType.TranslatedFrom, "16")]
    public void ClassificationIsExplicitAndDoesNotReplaceDistinctiveTitle(BookOnixAlternativeTitleType type, string code) {
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with {
            AlternativeTitles = [new() { Type = type, Title = "Alternative & title" }]
        }, BookOnixTests.TestSchema());
        var details = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "TitleDetail").ToArray();
        Assert.Equal(new[] { "01", code }, details.Select(e => e.Element(Ns + "TitleType")!.Value));
        Assert.Equal("Book & title", details[0].Descendants(Ns + "TitleText").Single().Value);
        Assert.Equal("Alternative & title", details[1].Descendants(Ns + "TitleText").Single().Value);
        Assert.Equal("01", details[1].Descendants(Ns + "TitleElementLevel").Single().Value);
        Assert.Null(details[1].Descendants(Ns + "TitleText").Single().Attribute("language"));
    }

    [Fact]
    public void AlternativeLanguagesSortingAndRepeatedTypesSurviveMessageCompositionWithoutMutatingBook() {
        var project = BookOnixTests.Project();
        project.Publication.AddTitle("selected", new() { Text = "Selected main title" });
        byte[] before = project.ToProjectBytes();
        var options = BookOnixTests.Options() with { TitleId = "selected", AlternativeTitles = [
            new() { Type = BookOnixAlternativeTitleType.OriginalLanguage, Title = "L’histoire", Subtitle = "Une étude", LanguageCode = "fre", TitleSorting = new() { Prefix = "L’" } },
            new() { Type = BookOnixAlternativeTitleType.OtherLanguage, Title = "歴史", LanguageCode = "jpn", TitleSorting = new() },
            new() { Type = BookOnixAlternativeTitleType.OtherLanguage, Title = "Historia", LanguageCode = "pol" }
        ] };
        var result = project.ExportOnix(options, BookOnixTests.TestSchema());
        var message = BookOnixMessage.Create([result], BookOnixTests.TestSchema());
        Assert.Equal(result.Bytes, message.Bytes);
        var titles = XDocument.Load(new MemoryStream(message.Bytes)).Descendants(Ns + "TitleDetail").ToArray();
        Assert.Equal(new[] { "01", "03", "06", "06" }, titles.Select(t => t.Element(Ns + "TitleType")!.Value));
        Assert.Equal("Selected main title", titles[0].Descendants(Ns + "TitleText").Single().Value);
        Assert.Equal("L’", titles[1].Descendants(Ns + "TitlePrefix").Single().Value);
        Assert.Equal("histoire", titles[1].Descendants(Ns + "TitleWithoutPrefix").Single().Value);
        Assert.Equal("Une étude", titles[1].Descendants(Ns + "Subtitle").Single().Value);
        Assert.All(titles[1].Descendants().Where(e => e.Name.LocalName is "TitlePrefix" or "TitleWithoutPrefix" or "Subtitle"),
            e => Assert.Equal("fre", (string?)e.Attribute("language")));
        Assert.Single(titles[2].Descendants(Ns + "NoPrefix"));
        Assert.Equal("jpn", (string?)titles[2].Descendants(Ns + "TitleWithoutPrefix").Single().Attribute("language"));
        Assert.Equal("Historia", titles[3].Descendants(Ns + "TitleText").Single().Value);
        Assert.Equal(project.Export().Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Theory]
    [InlineData("null-list")]
    [InlineData("null-item")]
    [InlineData("too-many")]
    [InlineData("type")]
    [InlineData("null-title")]
    [InlineData("blank-title")]
    [InlineData("oversized-title")]
    [InlineData("blank-subtitle")]
    [InlineData("language")]
    [InlineData("prefix")]
    public void InvalidAssertionsFailBeforeChangingTheProject(string invalid) {
        var title = new BookOnixAlternativeTitle { Type = BookOnixAlternativeTitleType.Former, Title = "Former title" };
        IReadOnlyList<BookOnixAlternativeTitle> titles = invalid switch {
            "null-list" => null!, "null-item" => [null!], "too-many" => Enumerable.Repeat(title, 33).ToArray(),
            "type" => [title with { Type = (BookOnixAlternativeTitleType)999 }],
            "null-title" => [title with { Title = null! }], "blank-title" => [title with { Title = " " }],
            "oversized-title" => [title with { Title = new string('x', 4097) }],
            "blank-subtitle" => [title with { Subtitle = " " }], "language" => [title with { LanguageCode = "en" }],
            "prefix" => [title with { TitleSorting = new() { Prefix = "The " } }], _ => throw new InvalidOperationException()
        };
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with {
            AlternativeTitles = titles }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

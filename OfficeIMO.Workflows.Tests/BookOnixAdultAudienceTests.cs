using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixAdultAudienceTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;

    [Theory]
    [InlineData(BookOnixAdultAudienceRating.Unrated, "00")]
    [InlineData(BookOnixAdultAudienceRating.AnyAdultAudience, "01")]
    [InlineData(BookOnixAdultAudienceRating.ContentAdvice, "02")]
    [InlineData(BookOnixAdultAudienceRating.SexualContent, "03")]
    [InlineData(BookOnixAdultAudienceRating.Violence, "04")]
    [InlineData(BookOnixAdultAudienceRating.DrugsOrAlcohol, "05")]
    [InlineData(BookOnixAdultAudienceRating.OffensiveLanguage, "06")]
    [InlineData(BookOnixAdultAudienceRating.Intolerance, "07")]
    [InlineData(BookOnixAdultAudienceRating.Abuse, "08")]
    [InlineData(BookOnixAdultAudienceRating.SelfHarm, "09")]
    [InlineData(BookOnixAdultAudienceRating.AnimalCruelty, "10")]
    [InlineData(BookOnixAdultAudienceRating.Illness, "11")]
    [InlineData(BookOnixAdultAudienceRating.DeathAndGrief, "12")]
    [InlineData(BookOnixAdultAudienceRating.Suicide, "13")]
    public void RatingsUseList203CodesWithinTheAdultRatingScheme(BookOnixAdultAudienceRating rating, string code) {
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Audience = new() {
            Categories = [new(BookOnixAudienceType.GeneralAdult)], AdultRatings = [new(rating, true)]
        } }, BookOnixTests.TestSchema());
        var entries = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "Audience").ToArray();
        Assert.Equal(new[] { "01", "22" }, entries.Select(e => e.Element(Ns + "AudienceCodeType")!.Value));
        Assert.Equal(code, entries[1].Element(Ns + "AudienceCodeValue")!.Value);
        Assert.NotNull(entries[1].Element(Ns + "MainAudience"));
    }

    [Fact]
    public void ContentAdviceCanBeCombinedWithTranslationsWithoutChangingThePublication() {
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { Audience = new() {
            Categories = [new(BookOnixAudienceType.GeneralAdult, true)],
            AdultRatings = Enum.GetValues<BookOnixAdultAudienceRating>().Where(rating => rating >= BookOnixAdultAudienceRating.ContentAdvice)
                .Select(rating => new BookOnixAdultAudience(rating, rating == BookOnixAdultAudienceRating.ContentAdvice) {
                    Headings = rating == BookOnixAdultAudienceRating.ContentAdvice ? [new("Publisher advice & context", "eng"), new("Informacja wydawcy", "pol")] : []
                }).ToArray()
        } }, BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        var ratings = xml.Descendants(Ns + "Audience").Where(e => e.Element(Ns + "AudienceCodeType")!.Value == "22").ToArray();
        Assert.Equal(new[] { "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13" }, ratings.Select(e => e.Element(Ns + "AudienceCodeValue")!.Value));
        Assert.Equal(new[] { "eng", "pol" }, ratings[0].Elements(Ns + "AudienceHeadingText").Select(e => (string?)e.Attribute("language")));
        Assert.Equal(2, xml.Descendants(Ns + "MainAudience").Count());
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
        Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
    }

    [Theory]
    [InlineData("missing-adult")]
    [InlineData("duplicate")]
    [InlineData("unrated-advice")]
    [InlineData("any-adult-advice")]
    [InlineData("unrated-any")]
    [InlineData("multiple-main")]
    [InlineData("unknown")]
    [InlineData("null-list")]
    [InlineData("null-entry")]
    [InlineData("count")]
    [InlineData("heading-language")]
    public void InvalidOrContradictoryAssertionsDoNotMutateTheProject(string kind) {
        IReadOnlyList<BookOnixAdultAudience> ratings = kind switch {
            "duplicate" => [new(BookOnixAdultAudienceRating.ContentAdvice), new(BookOnixAdultAudienceRating.ContentAdvice)],
            "unrated-advice" => [new(BookOnixAdultAudienceRating.Unrated), new(BookOnixAdultAudienceRating.ContentAdvice)],
            "any-adult-advice" => [new(BookOnixAdultAudienceRating.AnyAdultAudience), new(BookOnixAdultAudienceRating.ContentAdvice)],
            "unrated-any" => [new(BookOnixAdultAudienceRating.Unrated), new(BookOnixAdultAudienceRating.AnyAdultAudience)],
            "multiple-main" => [new(BookOnixAdultAudienceRating.Violence, true), new(BookOnixAdultAudienceRating.OffensiveLanguage, true)],
            "unknown" => [new((BookOnixAdultAudienceRating)99)],
            "null-list" => null!,
            "null-entry" => [null!],
            "count" => Enumerable.Repeat(new BookOnixAdultAudience(BookOnixAdultAudienceRating.ContentAdvice), 15).ToArray(),
            "heading-language" => [new(BookOnixAdultAudienceRating.ContentAdvice) { Headings = [new("Advice", "eng"), new("Informacja")] }],
            _ => [new(BookOnixAdultAudienceRating.ContentAdvice)]
        };
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Audience = new() {
            Categories = kind == "missing-adult" ? [new(BookOnixAudienceType.Children)] : [new(BookOnixAudienceType.GeneralAdult)],
            AdultRatings = ratings
        } }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

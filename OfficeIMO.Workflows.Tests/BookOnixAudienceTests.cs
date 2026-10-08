using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixAudienceTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixAudienceMetadata Audience() => new() { Categories = [new(BookOnixAudienceType.Children, true), new(BookOnixAudienceType.Teenage)],
        AgeRanges = [new(BookOnixAgeRangeType.InterestYears, 10, 14), new(BookOnixAgeRangeType.ReadingYears, 9, 12)],
        Descriptions = [new("Readers of adventure & exploration", "eng"), new("Czytelnicy przygód", "pol")] };

    [Fact]
    public void CategoriesAndDescriptionsRemainExplicitAndDoNotChangeThePublication() {
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { Audience = Audience() }, BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        Assert.Equal(new[] { "02", "03" }, xml.Descendants(Ns + "AudienceCodeValue").Select(e => e.Value));
        Assert.All(xml.Descendants(Ns + "AudienceCodeType"), e => Assert.Equal("01", e.Value));
        Assert.Single(xml.Descendants(Ns + "MainAudience"));
        Assert.Equal(new[] { "17", "18" }, xml.Descendants(Ns + "AudienceRangeQualifier").Select(e => e.Value));
        var descriptions = xml.Descendants(Ns + "AudienceDescription").ToArray();
        Assert.Equal(new[] { "eng", "pol" }, descriptions.Select(e => (string?)e.Attribute("language")));
        Assert.All(descriptions, e => Assert.Equal("06", (string?)e.Attribute("textformat")));
        Assert.Equal("Readers of adventure & exploration", descriptions[0].Value);
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void EverySupportedCategoryUsesItsList28Code() {
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Audience = new() {
            Categories = Enum.GetValues<BookOnixAudienceType>().Select(type => new BookOnixAudience(type)).ToArray()
        } }, BookOnixTests.TestSchema());
        Assert.Equal(new[] { "01", "02", "03", "04", "05", "06", "07", "08", "09", "11", "12", "13", "14" },
            XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "AudienceCodeValue").Select(e => e.Value));
    }

    [Theory]
    [InlineData(BookOnixAgeRangeType.InterestYears, 8, 12, "17", "03,04", "8,12")]
    [InlineData(BookOnixAgeRangeType.InterestYears, 8, 8, "17", "01", "8")]
    [InlineData(BookOnixAgeRangeType.InterestYears, 8, null, "17", "03", "8")]
    [InlineData(BookOnixAgeRangeType.InterestYears, null, 8, "17", "04", "8")]
    [InlineData(BookOnixAgeRangeType.InterestMonths, 0, 0, "16", "01", "0")]
    [InlineData(BookOnixAgeRangeType.InterestMonths, 36, 42, "16", "03,04", "36,42")]
    [InlineData(BookOnixAgeRangeType.InterestMonths, null, 36, "16", "04", "36")]
    [InlineData(BookOnixAgeRangeType.ReadingYears, 5, null, "18", "03", "5")]
    public void RangePrecisionPreservesExactOpenAndClosedBounds(BookOnixAgeRangeType type, int? minimum, int? maximum,
        string qualifier, string precision, string values) {
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Audience = new() {
            AgeRanges = [new(type, minimum, maximum)] } }, BookOnixTests.TestSchema());
        var range = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "AudienceRange").Single();
        Assert.Equal(qualifier, range.Element(Ns + "AudienceRangeQualifier")!.Value);
        Assert.Equal(precision.Split(','), range.Elements(Ns + "AudienceRangePrecision").Select(e => e.Value));
        Assert.Equal(values.Split(','), range.Elements(Ns + "AudienceRangeValue").Select(e => e.Value));
    }

    [Theory]
    [InlineData("empty")]
    [InlineData("category")]
    [InlineData("duplicate-category")]
    [InlineData("multiple-main")]
    [InlineData("range-type")]
    [InlineData("no-bounds")]
    [InlineData("negative")]
    [InlineData("reversed")]
    [InlineData("month-first")]
    [InlineData("month-second")]
    [InlineData("month-only-upper")]
    [InlineData("month-exact")]
    [InlineData("duplicate-range")]
    [InlineData("mixed-interest-units")]
    [InlineData("language")]
    [InlineData("duplicate-language")]
    [InlineData("description")]
    [InlineData("description-count")]
    public void InvalidAssertionsFailBeforeMutation(string kind) {
        var audience = kind switch {
            "empty" => new BookOnixAudienceMetadata(),
            "category" => Audience() with { Categories = [new((BookOnixAudienceType)99)] },
            "duplicate-category" => Audience() with { Categories = [new(BookOnixAudienceType.Children), new(BookOnixAudienceType.Children)] },
            "multiple-main" => Audience() with { Categories = [new(BookOnixAudienceType.Children, true), new(BookOnixAudienceType.Teenage, true)] },
            "range-type" => Audience() with { AgeRanges = [new((BookOnixAgeRangeType)99, 1)] },
            "no-bounds" => Audience() with { AgeRanges = [new(BookOnixAgeRangeType.InterestYears)] },
            "negative" => Audience() with { AgeRanges = [new(BookOnixAgeRangeType.InterestYears, -1)] },
            "reversed" => Audience() with { AgeRanges = [new(BookOnixAgeRangeType.InterestYears, 8, 3)] },
            "month-first" => Audience() with { AgeRanges = [new(BookOnixAgeRangeType.InterestMonths, 37, 42)] },
            "month-second" => Audience() with { AgeRanges = [new(BookOnixAgeRangeType.InterestMonths, 36, 43)] },
            "month-only-upper" => Audience() with { AgeRanges = [new(BookOnixAgeRangeType.InterestMonths, null, 42)] },
            "month-exact" => Audience() with { AgeRanges = [new(BookOnixAgeRangeType.InterestMonths, 37, 37)] },
            "duplicate-range" => Audience() with { AgeRanges = [new(BookOnixAgeRangeType.InterestYears, 1), new(BookOnixAgeRangeType.InterestYears, 2)] },
            "mixed-interest-units" => Audience() with { AgeRanges = [new(BookOnixAgeRangeType.InterestYears, 1), new(BookOnixAgeRangeType.InterestMonths, 12)] },
            "language" => Audience() with { Descriptions = [new("Audience", "en")] },
            "duplicate-language" => Audience() with { Descriptions = [new("One"), new("Two")] },
            "description" => Audience() with { Descriptions = [new(" ")] },
            _ => Audience() with { Descriptions = Enumerable.Repeat(new BookOnixAudienceDescription("Text"), 17).ToArray() }
        };
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Audience = audience }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

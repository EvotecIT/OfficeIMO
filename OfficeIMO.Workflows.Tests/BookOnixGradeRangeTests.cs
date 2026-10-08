using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixGradeRangeTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixExportResult Export(BookOnixGradeRange range) => BookOnixTests.Project().ExportOnix(
        BookOnixTests.Options() with { Audience = new() { GradeRanges = [range] } }, BookOnixTests.TestSchema());

    [Theory]
    [InlineData(BookOnixGradeSystem.UnitedStates, "11")]
    [InlineData(BookOnixGradeSystem.CanadaExcludingQuebec, "26")]
    [InlineData(BookOnixGradeSystem.China, "29")]
    public void EveryGradeRetainsItsCodeWithinEachSystem(BookOnixGradeSystem system, string qualifier) {
        foreach (var grade in Enum.GetValues<BookOnixGrade>()) {
            var result = Export(new(system, grade, grade));
            var range = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "AudienceRange").Single();
            Assert.Equal(qualifier, range.Element(Ns + "AudienceRangeQualifier")!.Value);
            Assert.Equal("01", range.Element(Ns + "AudienceRangePrecision")!.Value);
            string expected = grade == BookOnixGrade.Preschool ? "P" : grade == BookOnixGrade.Kindergarten ? "K" : ((int)grade).ToString(System.Globalization.CultureInfo.InvariantCulture);
            Assert.Equal(expected, range.Element(Ns + "AudienceRangeValue")!.Value);
        }
    }

    [Theory]
    [InlineData(BookOnixGrade.Preschool, BookOnixGrade.Kindergarten, "03,04", "P,K")]
    [InlineData(BookOnixGrade.Kindergarten, BookOnixGrade.Grade2, "03,04", "K,2")]
    [InlineData(BookOnixGrade.Grade9, BookOnixGrade.Grade12, "03,04", "9,12")]
    [InlineData(BookOnixGrade.Grade13, null, "03", "13")]
    [InlineData(null, BookOnixGrade.Kindergarten, "04", "K")]
    public void PrecisionUsesGradeOrderRatherThanLexicalOrAgeOrder(BookOnixGrade? minimum, BookOnixGrade? maximum, string precision, string values) {
        var result = Export(new(BookOnixGradeSystem.UnitedStates, minimum, maximum));
        var range = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "AudienceRange").Single();
        Assert.Equal(precision.Split(','), range.Elements(Ns + "AudienceRangePrecision").Select(e => e.Value));
        Assert.Equal(values.Split(','), range.Elements(Ns + "AudienceRangeValue").Select(e => e.Value));
    }

    [Fact]
    public void GradesCoexistWithAgeAndCategoriesWithoutChangingTheBookOrInferringAges() {
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { Audience = new() {
            Categories = [new(BookOnixAudienceType.PrimaryEducation)],
            AgeRanges = [new(BookOnixAgeRangeType.InterestYears, 6, 8)],
            GradeRanges = Enum.GetValues<BookOnixGradeSystem>().Select(system => new BookOnixGradeRange(system, BookOnixGrade.Grade1, BookOnixGrade.Grade3)).ToArray(),
            Descriptions = [new("Publisher-supplied educational levels", "eng")]
        } }, BookOnixTests.TestSchema());
        Assert.Equal(new[] { "17", "11", "26", "29" }, XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "AudienceRangeQualifier").Select(e => e.Value));
        Assert.Equal(before, project.ToProjectBytes());
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Single(XDocument.Load(new MemoryStream(Export(new(BookOnixGradeSystem.China, BookOnixGrade.Kindergarten)).Bytes)).Descendants(Ns + "AudienceRange"));
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("system")]
    [InlineData("lower")]
    [InlineData("upper")]
    [InlineData("reversed-preschool")]
    [InlineData("reversed-number")]
    [InlineData("duplicate")]
    [InlineData("count")]
    [InlineData("null-item")]
    [InlineData("null-list")]
    public void InvalidRangesRejectWithoutMutation(string scenario) {
        BookOnixGradeRange valid = new(BookOnixGradeSystem.UnitedStates, BookOnixGrade.Grade1, BookOnixGrade.Grade4);
        IReadOnlyList<BookOnixGradeRange> ranges = scenario switch {
            "missing" => [new(BookOnixGradeSystem.UnitedStates)],
            "system" => [valid with { System = (BookOnixGradeSystem)99 }],
            "lower" => [valid with { Minimum = (BookOnixGrade)(-2) }],
            "upper" => [valid with { Maximum = (BookOnixGrade)18 }],
            "reversed-preschool" => [valid with { Minimum = BookOnixGrade.Kindergarten, Maximum = BookOnixGrade.Preschool }],
            "reversed-number" => [valid with { Minimum = BookOnixGrade.Grade12, Maximum = BookOnixGrade.Grade9 }],
            "duplicate" => [valid, valid with { Minimum = BookOnixGrade.Grade2 }],
            "count" => [valid, valid, valid, valid],
            "null-item" => [null!], _ => null!
        };
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with {
            Audience = new() { GradeRanges = ranges } }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

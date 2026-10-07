using System.Globalization;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixReviewRatingTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixCollateralText Item(BookOnixReviewRating rating) => new() {
        Type = BookOnixTextType.ReviewQuote, Audiences = [BookOnixContentAudience.EndCustomers],
        Texts = [new("A publisher-supplied review.")], ReviewRating = rating
    };

    [Theory]
    [InlineData("0", null)]
    [InlineData("0", 1)]
    [InlineData("4.50", 5)]
    [InlineData("100", 100)]
    [InlineData("7.25", null)]
    [InlineData("79228162514264337593543950335", null)]
    public void ScoresRetainDecimalPrecisionAndOptionalLimits(string value, int? limit) {
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with {
            CollateralTexts = [Item(new(decimal.Parse(value, CultureInfo.InvariantCulture), limit))]
        }, BookOnixTests.TestSchema());
        var rating = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "ReviewRating").Single();
        Assert.Equal(value, rating.Element(Ns + "Rating")!.Value);
        Assert.Equal(limit?.ToString(CultureInfo.InvariantCulture), rating.Element(Ns + "RatingLimit")?.Value);
        Assert.Empty(rating.Elements(Ns + "RatingUnits"));
        Assert.Equal(before, project.ToProjectBytes());
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
    }

    [Fact]
    public void ReviewRolesAndSingleUnspecifiedUnitAreSupported() {
        var items = new[] { BookOnixTextType.ReviewQuote, BookOnixTextType.PreviousEditionReview, BookOnixTextType.PreviousWorkReview }
            .Select(type => Item(new(5, 5) { Units = [new(new string('A', 50))] }) with { Type = type }).ToArray();
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { CollateralTexts = items }, BookOnixTests.TestSchema());
        var units = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "RatingUnits").ToArray();
        Assert.Equal(3, units.Length);
        Assert.All(units, unit => { Assert.Equal(new string('A', 50), unit.Value); Assert.Null(unit.Attribute("language")); });
    }

    [Fact]
    public void TranslatedUnitsRemainPlainTextAndNumbersIgnoreCurrentCulture() {
        var previous = CultureInfo.CurrentCulture;
        try {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("pl-PL");
            var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { CollateralTexts = [
                Item(new(4.5m, 5) { Units = [new("<stars> & points", "eng"), new("gwiazdki", "pol")] }) with {
                    Authors = ["Example Reviewer"], SourceTitles = [new("Example Journal")]
                }]
            }, BookOnixTests.TestSchema());
            var content = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "TextContent").Single();
            Assert.Equal(new[] { "SequenceNumber", "TextType", "ContentAudience", "Text", "ReviewRating", "TextAuthor", "SourceTitle" },
                content.Elements().Select(e => e.Name.LocalName));
            var rating = content.Element(Ns + "ReviewRating")!;
            Assert.Equal("4.5", rating.Element(Ns + "Rating")!.Value);
            var units = rating.Elements(Ns + "RatingUnits").ToArray();
            Assert.Equal(new[] { "eng", "pol" }, units.Select(e => (string?)e.Attribute("language")));
            Assert.Equal("<stars> & points", units[0].Value);
            Assert.All(units, unit => Assert.Empty(unit.Elements()));
        } finally { CultureInfo.CurrentCulture = previous; }
    }

    [Theory]
    [InlineData("negative")]
    [InlineData("zero-limit")]
    [InlineData("negative-limit")]
    [InlineData("over-limit")]
    [InlineData("description")]
    [InlineData("endorsement")]
    [InlineData("null-units")]
    [InlineData("null-entry")]
    [InlineData("unit-count")]
    [InlineData("blank-unit")]
    [InlineData("long-unit")]
    [InlineData("language")]
    [InlineData("missing-language")]
    [InlineData("duplicate-language")]
    public void InvalidRatingFailsWithoutMutation(string kind) {
        var rating = kind switch {
            "negative" => new BookOnixReviewRating(-0.1m), "zero-limit" => new(0, 0),
            "negative-limit" => new(0, -1), "over-limit" => new(5.1m, 5),
            "null-units" => new(1) { Units = null! }, "null-entry" => new(1) { Units = [null!] },
            "unit-count" => new(1) { Units = Enumerable.Repeat(new BookOnixRatingUnit("stars"), 17).ToArray() },
            "blank-unit" => new(1) { Units = [new(" ")] }, "long-unit" => new(1) { Units = [new(new string('A', 51))] },
            "language" => new(1) { Units = [new("stars", "en-US")] },
            "missing-language" => new(1) { Units = [new("stars", "eng"), new("gwiazdki")] },
            "duplicate-language" => new(1) { Units = [new("stars", "eng"), new("points", "eng")] },
            _ => new(1)
        };
        var item = Item(rating) with { Type = kind == "description" ? BookOnixTextType.Description :
            kind == "endorsement" ? BookOnixTextType.Endorsement : BookOnixTextType.ReviewQuote };
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { CollateralTexts = [item] }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void RatingUnitsShareTheCollateralTextBudget() {
        var item = Item(new(1) { Units = [new("stars")] }) with { Texts = [new(new string('A', 65536))] };
        Assert.Throws<ArgumentException>(() => BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with {
            CollateralTexts = Enumerable.Repeat(item, 8).ToArray()
        }, BookOnixTests.TestSchema()));
    }
}

using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixLicenseTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixLicense License() => new() { Names = [new("Publisher terms")] };
    private static BookOnixCollateralText Item(params BookOnixLicense[] licenses) => new() {
        Type = BookOnixTextType.Excerpt, Audiences = [BookOnixContentAudience.EndCustomers],
        Texts = [new("A supplied excerpt.")], Licenses = licenses
    };

    [Theory]
    [InlineData(BookOnixLicenseExpressionType.HumanReadable, "01")]
    [InlineData(BookOnixLicenseExpressionType.ProfessionalReadable, "02")]
    [InlineData(BookOnixLicenseExpressionType.AdditionalHumanReadable, "03")]
    [InlineData(BookOnixLicenseExpressionType.AdditionalProfessionalReadable, "04")]
    [InlineData(BookOnixLicenseExpressionType.OnixPl, "10")]
    [InlineData(BookOnixLicenseExpressionType.Odrl, "20")]
    [InlineData(BookOnixLicenseExpressionType.AdditionalOdrl, "21")]
    public void ExpressionTypesUseList218AndPreserveTheLink(BookOnixLicenseExpressionType type, string code) {
        const string link = "https://example.org/terms?a=1&b=2#text";
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with {
            CollateralTexts = [Item(License() with { Expressions = [new(type, link)] })]
        }, BookOnixTests.TestSchema());
        var expression = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "EpubLicenseExpression").Single();
        Assert.Equal(code, expression.Element(Ns + "EpubLicenseExpressionType")!.Value);
        Assert.Equal(link, expression.Element(Ns + "EpubLicenseExpressionLink")!.Value);
        Assert.Null(expression.Element(Ns + "EpubLicenseExpressionTypeName"));
    }

    [Fact]
    public void MultipleDatedLicensesKeepNamesAndEveryBoundaryWithoutChangingThePublication() {
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { CollateralTexts = [Item(
            License() with { Names = [new("Original <terms> & text", "eng"), new("Warunki", "pol")],
                ValidFrom = new(2026, 1, 1), ValidUntil = new(2026, 12, 31) },
            License() with { ValidFrom = new(2027, 1, 1), ValidUntil = new(2027, 12, 31) }) with {
                SourceLinks = ["https://example.org/excerpt"], PublishedOn = new(2026, 1, 1)
            }]
        }, BookOnixTests.TestSchema());
        var content = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "TextContent").Single();
        var licenses = content.Elements(Ns + "EpubLicense").ToArray();
        Assert.Equal(2, licenses.Length);
        Assert.Equal(new[] { "eng", "pol" }, licenses[0].Elements(Ns + "EpubLicenseName").Select(e => (string?)e.Attribute("language")));
        Assert.Equal("Original <terms> & text", licenses[0].Element(Ns + "EpubLicenseName")!.Value);
        Assert.All(licenses, license => Assert.Equal(new[] { "14", "15" }, license.Descendants(Ns + "EpubLicenseDateRole").Select(e => e.Value)));
        Assert.Equal(new[] { "20260101", "20261231", "20270101", "20271231" }, licenses.SelectMany(e => e.Descendants(Ns + "Date")).Select(e => e.Value));
        Assert.Equal("TextSourceLink", licenses[0].ElementsBeforeSelf().Last().Name.LocalName);
        Assert.Equal("ContentDate", licenses[1].ElementsAfterSelf().First().Name.LocalName);
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
        Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
    }

    [Fact]
    public void NamesOnlyOpenDatesAndSingleDayPeriodsArePreserved() {
        var day = new DateOnly(2026, 10, 6);
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { CollateralTexts = [Item(
            License() with { Names = [new(new string('A', 100))] }, License() with { ValidFrom = day },
            License() with { ValidUntil = day }, License() with { ValidFrom = day, ValidUntil = day })]
        }, BookOnixTests.TestSchema());
        var licenses = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "EpubLicense").ToArray();
        Assert.Empty(licenses[0].Elements(Ns + "EpubLicenseDate"));
        Assert.All(licenses, e => Assert.Empty(e.Elements(Ns + "EpubLicenseExpression")));
        Assert.Equal(new[] { 0, 1, 1, 2 }, licenses.Select(e => e.Elements(Ns + "EpubLicenseDate").Count()));
        Assert.Equal(new[] { "14", "15", "14", "15" }, licenses.SelectMany(e => e.Descendants(Ns + "EpubLicenseDateRole")).Select(e => e.Value));
    }

    [Theory]
    [InlineData("no-names")]
    [InlineData("null-names")]
    [InlineData("null-name")]
    [InlineData("name-count")]
    [InlineData("blank-name")]
    [InlineData("long-name")]
    [InlineData("language")]
    [InlineData("missing-language")]
    [InlineData("duplicate-language")]
    [InlineData("null-expressions")]
    [InlineData("null-expression")]
    [InlineData("expression-count")]
    [InlineData("duplicate-expression")]
    [InlineData("unknown-type")]
    [InlineData("link-scheme")]
    [InlineData("link-credentials")]
    [InlineData("link-length")]
    [InlineData("dates")]
    [InlineData("null-list")]
    [InlineData("null-license")]
    [InlineData("license-count")]
    public void InvalidLicenseMetadataFailsWithoutMutation(string kind) {
        var expression = new BookOnixLicenseExpression(BookOnixLicenseExpressionType.HumanReadable, "https://example.org/terms");
        var license = kind switch {
            "no-names" => License() with { Names = [] }, "null-names" => License() with { Names = null! },
            "null-name" => License() with { Names = [null!] },
            "name-count" => License() with { Names = Enumerable.Repeat(new BookOnixLicenseName("Terms"), 17).ToArray() },
            "blank-name" => License() with { Names = [new(" ")] }, "long-name" => License() with { Names = [new(new string('A', 101))] },
            "language" => License() with { Names = [new("Terms", "en")] },
            "missing-language" => License() with { Names = [new("Terms", "eng"), new("Warunki")] },
            "duplicate-language" => License() with { Names = [new("Terms", "eng"), new("Other", "eng")] },
            "null-expressions" => License() with { Expressions = null! }, "null-expression" => License() with { Expressions = [null!] },
            "expression-count" => License() with { Expressions = Enumerable.Repeat(expression, 17).ToArray() },
            "duplicate-expression" => License() with { Expressions = [expression, expression] },
            "unknown-type" => License() with { Expressions = [expression with { Type = (BookOnixLicenseExpressionType)99 }] },
            "link-scheme" => License() with { Expressions = [expression with { Link = "file:///tmp/terms" }] },
            "link-credentials" => License() with { Expressions = [expression with { Link = "https://user:secret@example.org/terms" }] },
            "link-length" => License() with { Expressions = [expression with { Link = "https://example.org/" + new string('A', 4096) }] },
            "dates" => License() with { ValidFrom = new(2027, 1, 1), ValidUntil = new(2026, 1, 1) },
            _ => License()
        };
        var item = Item(license) with { Licenses = kind switch {
            "null-list" => null!, "null-license" => [null!],
            "license-count" => Enumerable.Repeat(license, 17).ToArray(), _ => [license]
        } };
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { CollateralTexts = [item] }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NamesAndExpressionLinksConsumeTheCollateralBudget(bool link) {
        var license = License() with { Names = [new("A")], Expressions = link ? [new(BookOnixLicenseExpressionType.HumanReadable, "https://example.org/terms")] : [] };
        var item = Item(license) with { Texts = [new(new string('A', link ? 65535 : 65536))] };
        Assert.Throws<ArgumentException>(() => BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with {
            CollateralTexts = Enumerable.Repeat(item, 8).ToArray()
        }, BookOnixTests.TestSchema()));
    }
}

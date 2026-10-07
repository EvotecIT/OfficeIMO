using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixAudienceCodeTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;

    [Fact]
    public void AdditionalSchemesPreserveCodesAndCoexistWithGeneralCategoriesAndRanges() {
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var codes = Enum.GetValues<BookOnixAudienceScheme>().Select(scheme => new BookOnixAudienceCode(scheme,
            scheme == BookOnixAudienceScheme.Cefr ? "B2" : scheme == BookOnixAudienceScheme.IntendedLanguage ? "pol" : "01") {
            SchemeName = scheme == BookOnixAudienceScheme.Proprietary ? "House & recipient scheme" : null, IsMain = true
        }).ToArray();
        var result = project.ExportOnix(BookOnixTests.Options() with { Audience = new() {
            Categories = [new(BookOnixAudienceType.Children, true)], Codes = codes,
            AgeRanges = [new(BookOnixAgeRangeType.InterestYears, 8, 12)]
        } }, BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        var audiences = xml.Descendants(Ns + "Audience").ToArray();
        Assert.Equal(new[] { "01", "02", "06", "07", "08", "09", "11", "15", "16", "17", "18", "21", "23", "27", "29", "30" },
            audiences.Select(e => e.Element(Ns + "AudienceCodeType")!.Value));
        Assert.All(audiences, e => Assert.NotNull(e.Element(Ns + "MainAudience")));
        Assert.Equal(codes.Select(code => code.Value), audiences.Skip(1).Select(e => e.Element(Ns + "AudienceCodeValue")!.Value));
        Assert.Equal("House & recipient scheme", Assert.Single(xml.Descendants(Ns + "AudienceCodeTypeName")).Value);
        Assert.Single(xml.Descendants(Ns + "AudienceRange"));
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void NamedProprietarySchemesRetainDistinctIdentitiesAndNeedNoGenericAudience() {
        var result = Export([
            new(BookOnixAudienceScheme.Proprietary, "01") { SchemeName = "Publisher", IsMain = true },
            new(BookOnixAudienceScheme.Proprietary, "02") { SchemeName = "Publisher" },
            new(BookOnixAudienceScheme.Proprietary, "01") { SchemeName = "Recipient" }
        ]);
        var audiences = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "Audience").ToArray();
        Assert.Equal(new[] { "Publisher", "Publisher", "Recipient" }, audiences.Select(e => e.Element(Ns + "AudienceCodeTypeName")!.Value));
        Assert.Equal(new[] { "01", "02", "01" }, audiences.Select(e => e.Element(Ns + "AudienceCodeValue")!.Value));
        Assert.Single(audiences.SelectMany(e => e.Elements(Ns + "MainAudience")));
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("A2")]
    [InlineData("B1")]
    [InlineData("B2")]
    [InlineData("C1")]
    [InlineData("C2")]
    public void CefrUsesThePublishedLevelCodes(string value) {
        var result = Export([new(BookOnixAudienceScheme.Cefr, value)]);
        Assert.Equal(value, XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "AudienceCodeValue").Single().Value);
    }

    [Theory]
    [InlineData("missing-name")]
    [InlineData("unexpected-name")]
    [InlineData("name-whitespace")]
    [InlineData("empty-value")]
    [InlineData("value-whitespace")]
    [InlineData("unknown-scheme")]
    [InlineData("duplicate")]
    [InlineData("multiple-main")]
    [InlineData("proprietary-main")]
    [InlineData("cefr")]
    [InlineData("japanese-width")]
    [InlineData("japanese-digits")]
    [InlineData("language")]
    [InlineData("limit")]
    [InlineData("null-codes")]
    [InlineData("null-code")]
    public void InvalidOrAmbiguousAssertionsDoNotMutateProject(string kind) {
        IReadOnlyList<BookOnixAudienceCode> codes = kind switch {
            "missing-name" => [new(BookOnixAudienceScheme.Proprietary, "A")],
            "unexpected-name" => [new(BookOnixAudienceScheme.Electre, "A") { SchemeName = "Other" }],
            "name-whitespace" => [new(BookOnixAudienceScheme.Proprietary, "A") { SchemeName = " House" }],
            "empty-value" => [new(BookOnixAudienceScheme.Btlf, " ")],
            "value-whitespace" => [new(BookOnixAudienceScheme.Btlf, "01 ")],
            "unknown-scheme" => [new((BookOnixAudienceScheme)999, "A")],
            "duplicate" => [new(BookOnixAudienceScheme.Btlf, "01"), new(BookOnixAudienceScheme.Btlf, "01")],
            "multiple-main" => [new(BookOnixAudienceScheme.Btlf, "01") { IsMain = true }, new(BookOnixAudienceScheme.Btlf, "02") { IsMain = true }],
            "proprietary-main" => [new(BookOnixAudienceScheme.Proprietary, "01") { SchemeName = "One", IsMain = true },
                new(BookOnixAudienceScheme.Proprietary, "01") { SchemeName = "Two", IsMain = true }],
            "cefr" => [new(BookOnixAudienceScheme.Cefr, "B3")],
            "japanese-width" => [new(BookOnixAudienceScheme.JapaneseChildren, "1")],
            "japanese-digits" => [new(BookOnixAudienceScheme.JapaneseChildren, "１２")],
            "language" => [new(BookOnixAudienceScheme.IntendedLanguage, "pl")],
            "limit" => Enumerable.Range(0, 65).Select(value => new BookOnixAudienceCode(BookOnixAudienceScheme.Btlf, value.ToString())).ToArray(),
            "null-codes" => null!,
            _ => [null!]
        };
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with {
            Audience = new() { Codes = codes }
        }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    private static BookOnixExportResult Export(IReadOnlyList<BookOnixAudienceCode> codes) =>
        BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Audience = new() { Codes = codes } }, BookOnixTests.TestSchema());
}

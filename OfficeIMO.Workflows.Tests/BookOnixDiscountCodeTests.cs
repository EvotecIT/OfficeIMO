using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixDiscountCodeTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixDiscountCode Code() => new() { Scheme = BookOnixDiscountScheme.ProprietaryDiscount,
        SchemeName = "Example & partners", Code = "trade-A" };
    private static BookOnixExportOptions Options(params BookOnixDiscountCode[] codes) {
        var territory = new BookOnixTerritory { Countries = ["GB"] };
        return BookOnixTests.Options() with { Commercial = new() {
            SalesRights = [new(BookOnixSalesRightsKind.Exclusive, territory)], Supplies = [new() {
                Territory = territory, SupplierName = "Example Supplier", SupplierRole = BookOnixSupplierRole.PublisherToResellers,
                Availability = BookOnixAvailability.Available, Prices = [new() {
                    Kind = BookOnixPriceKind.RecommendedExcludingTax, Amount = 10m, CurrencyCode = "GBP", DiscountCodes = codes,
                    Discounts = [new() { Kind = BookOnixDiscountKind.Rising, Percent = 10m }]
                }]
            }]
        } };
    }

    [Fact]
    public void CodesKeepSchemeNamesOrderAndValuesAlongsideNumericDiscounts() {
        var codes = Enum.GetValues<BookOnixDiscountScheme>().Select(scheme => new BookOnixDiscountCode {
            Scheme = scheme,
            SchemeName = scheme is BookOnixDiscountScheme.ProprietaryDiscount or BookOnixDiscountScheme.ProprietaryCommission ? "Example & partners" : null,
            Code = scheme switch {
                BookOnixDiscountScheme.BicDiscount or BookOnixDiscountScheme.BicCommission => "ABCDE12",
                BookOnixDiscountScheme.IsniDiscount => "0000000121032683-A1",
                _ => "trade & code"
            }
        }).ToArray();
        var result = BookOnixTests.Project().ExportOnix(Options(codes), BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        var elements = xml.Descendants(Ns + "DiscountCoded").ToArray();
        Assert.Equal(new[] { "01", "02", "03", "04", "05", "06", "07" }, elements.Select(e => e.Element(Ns + "DiscountCodeType")!.Value));
        Assert.Equal(codes.Select(c => c.Code), elements.Select(e => e.Element(Ns + "DiscountCode")!.Value));
        Assert.Equal(new[] { "Example & partners", "Example & partners" }, xml.Descendants(Ns + "DiscountCodeTypeName").Select(e => e.Value));
        Assert.Equal("10", xml.Descendants(Ns + "DiscountPercent").Single().Value);
        Assert.Equal("10", xml.Descendants(Ns + "PriceAmount").Single().Value);
    }

    [Theory]
    [InlineData("missing-name")]
    [InlineData("blank-name")]
    [InlineData("unexpected-name")]
    [InlineData("empty-code")]
    [InlineData("long-code")]
    [InlineData("bic")]
    [InlineData("isni")]
    [InlineData("scheme")]
    [InlineData("count")]
    public void InvalidSchemeDeclarationsAreRejectedWithoutMutation(string kind) {
        var code = kind switch {
            "missing-name" => Code() with { SchemeName = null },
            "blank-name" => Code() with { Scheme = BookOnixDiscountScheme.ProprietaryCommission, SchemeName = " " },
            "unexpected-name" => Code() with { Scheme = BookOnixDiscountScheme.GermanTerms },
            "empty-code" => Code() with { Code = " " },
            "long-code" => Code() with { Code = new string('A', 4097) },
            "bic" => Code() with { Scheme = BookOnixDiscountScheme.BicCommission, SchemeName = null, Code = "ABCDE" },
            "isni" => Code() with { Scheme = BookOnixDiscountScheme.IsniDiscount, SchemeName = null, Code = "0000000121032683" },
            "scheme" => Code() with { Scheme = (BookOnixDiscountScheme)99 },
            _ => Code()
        };
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        var options = Options(kind == "count" ? Enumerable.Repeat(code, 17).ToArray() : [code]);
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(options, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

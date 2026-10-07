using System.Globalization;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixTaxTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixTax Tax() => new() { Type = BookOnixTaxType.ValueAdded, RateCode = BookOnixTaxRateCode.Lower,
        RatePercent = 5.00m, TaxableAmount = 10.000m, Amount = 0.5000m, PricePartDescription = "Digital text & illustrations" };
    private static BookOnixPrice Price() => new() { Kind = BookOnixPriceKind.RecommendedIncludingTax, Amount = 10.50m,
        CurrencyCode = "PLN", Taxes = [Tax()] };
    private static BookOnixExportOptions Options(BookOnixPrice price) {
        var territory = new BookOnixTerritory { Countries = ["PL"] };
        return BookOnixTests.Options() with { Commercial = new() {
            SalesRights = [new(BookOnixSalesRightsKind.Exclusive, territory)], Supplies = [new() {
                Territory = territory, SupplierName = "Example", SupplierRole = BookOnixSupplierRole.PublisherToCustomers,
                Availability = BookOnixAvailability.Available, Prices = [price]
            }]
        } };
    }
    private static XDocument Export(BookOnixPrice price) => XDocument.Load(new MemoryStream(
        BookOnixTests.Project().ExportOnix(Options(price), BookOnixTests.TestSchema()).Bytes));

    [Fact]
    public void TaxComponentsKeepTheirOrderDecimalScaleAndCurrencyContext() {
        CultureInfo previous = CultureInfo.CurrentCulture;
        try {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("pl-PL");
            var xml = Export(Price() with { Amount = 11.00m, Taxes = [Tax(),
                new() { Type = BookOnixTaxType.Environmental, Amount = 0.500m, PricePartDescription = "Separate component" }] });
            var taxes = xml.Descendants(Ns + "Tax").ToArray();
            Assert.Equal(new[] { "01", "03" }, taxes.Select(e => e.Element(Ns + "TaxType")!.Value));
            Assert.Equal("5.00", taxes[0].Element(Ns + "TaxRatePercent")!.Value);
            Assert.Equal("10.000", taxes[0].Element(Ns + "TaxableAmount")!.Value);
            Assert.Equal("0.5000", taxes[0].Element(Ns + "TaxAmount")!.Value);
            Assert.Equal("Digital text & illustrations", taxes[0].Element(Ns + "PricePartDescription")!.Value);
            Assert.Null(taxes[1].Element(Ns + "TaxRatePercent"));
            Assert.Equal("PLN", taxes[0].Parent!.Element(Ns + "CurrencyCode")!.Value);
        } finally { CultureInfo.CurrentCulture = previous; }
    }

    [Fact]
    public void UnknownExemptAndZeroRatedHaveDifferentSerializedMeanings() {
        var unknown = Export(Price() with { Taxes = [] });
        Assert.Empty(unknown.Descendants(Ns + "Tax")); Assert.Empty(unknown.Descendants(Ns + "TaxExempt"));
        var exempt = Export(Price() with { Taxes = [], TaxExempt = true });
        Assert.Single(exempt.Descendants(Ns + "TaxExempt")); Assert.Empty(exempt.Descendants(Ns + "Tax"));
        var zero = Export(Price() with { Taxes = [Tax() with { RateCode = BookOnixTaxRateCode.Zero, RatePercent = 0, Amount = 0 }] });
        Assert.Equal("Z", zero.Descendants(Ns + "TaxRateCode").Single().Value);
        Assert.Empty(zero.Descendants(Ns + "TaxExempt"));
        var rateOnly = Export(Price() with { Taxes = [Tax() with { Amount = null, TaxableAmount = null }] });
        Assert.Empty(rateOnly.Descendants(Ns + "TaxAmount"));
    }

    [Theory]
    [InlineData("exclusive")]
    [InlineData("exempt")]
    [InlineData("missing")]
    [InlineData("negative-rate")]
    [InlineData("large-rate")]
    [InlineData("negative-amount")]
    [InlineData("zero-base")]
    [InlineData("large-base")]
    [InlineData("large-total")]
    [InlineData("component-total")]
    [InlineData("zero-code")]
    [InlineData("zero-percent")]
    [InlineData("type")]
    [InlineData("code")]
    [InlineData("count")]
    public void ContradictoryOrUnsupportedTaxesAreRejectedWithoutMutation(string kind) {
        var price = kind switch {
            "exclusive" => Price() with { Kind = BookOnixPriceKind.RecommendedExcludingTax },
            "exempt" => Price() with { TaxExempt = true },
            "missing" => Price() with { Taxes = [Tax() with { RatePercent = null, Amount = null }] },
            "negative-rate" => Price() with { Taxes = [Tax() with { RatePercent = -1 }] },
            "large-rate" => Price() with { Taxes = [Tax() with { RatePercent = 101 }] },
            "negative-amount" => Price() with { Taxes = [Tax() with { Amount = -1 }] },
            "zero-base" => Price() with { Taxes = [Tax() with { TaxableAmount = 0 }] },
            "large-base" => Price() with { Taxes = [Tax() with { TaxableAmount = 11 }] },
            "component-total" => Price() with { Taxes = [Tax() with { Amount = 1 }] },
            "large-total" => Price() with { Taxes = [Tax() with { TaxableAmount = null, Amount = 6 }, Tax() with { TaxableAmount = null, Amount = 5 }] },
            "zero-code" => Price() with { Taxes = [Tax() with { RateCode = BookOnixTaxRateCode.Zero }] },
            "zero-percent" => Price() with { Taxes = [Tax() with { RatePercent = 0 }] },
            "type" => Price() with { Taxes = [Tax() with { Type = (BookOnixTaxType)99 }] },
            "code" => Price() with { Taxes = [Tax() with { RateCode = (BookOnixTaxRateCode)99 }] },
            _ => Price() with { Taxes = Enumerable.Repeat(Tax(), 17).ToArray() }
        };
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(Options(price), BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

using System.Globalization;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixDiscountTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixDiscount Discount() => new() { Kind = BookOnixDiscountKind.Rising, Percent = 10.00m, Amount = 1.000m };
    private static BookOnixExportOptions Options(params BookOnixDiscount[] discounts) {
        var territory = new BookOnixTerritory { Countries = ["GB"] };
        return BookOnixTests.Options() with { Commercial = new() {
            SalesRights = [new(BookOnixSalesRightsKind.Exclusive, territory)], Supplies = [new() {
                Territory = territory, SupplierName = "Example Supplier", SupplierRole = BookOnixSupplierRole.PublisherToResellers,
                Availability = BookOnixAvailability.Available, Prices = [new() {
                    Kind = BookOnixPriceKind.RecommendedExcludingTax, Amount = 10m, CurrencyCode = "GBP", Discounts = discounts
                }]
            }]
        } };
    }

    [Fact]
    public void DiscountsPreserveTypeOrderQuantitiesAndDecimalScaleWithoutChangingPrice() {
        CultureInfo before = CultureInfo.CurrentCulture;
        try {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("pl-PL");
            var result = BookOnixTests.Project().ExportOnix(Options(
                Discount() with { MinimumQuantity = 1, MaximumQuantity = 9 },
                Discount() with { Kind = BookOnixDiscountKind.Progressive, Percent = 20.00m, Amount = null, MinimumQuantity = 10 },
                Discount() with { Kind = BookOnixDiscountKind.RisingCumulative, Percent = null },
                Discount() with { Kind = BookOnixDiscountKind.ProgressiveCumulative, Percent = 0, Amount = 0 }), BookOnixTests.TestSchema());
            var xml = XDocument.Load(new MemoryStream(result.Bytes));
            var discounts = xml.Descendants(Ns + "Discount").ToArray();
            Assert.Equal(new[] { "01", "03", "02", "04" }, discounts.Select(e => e.Element(Ns + "DiscountType")!.Value));
            Assert.Equal("10.00", discounts[0].Element(Ns + "DiscountPercent")!.Value);
            Assert.Equal("1.000", discounts[0].Element(Ns + "DiscountAmount")!.Value);
            Assert.Equal("1", discounts[0].Element(Ns + "Quantity")!.Value);
            Assert.Equal("9", discounts[0].Element(Ns + "ToQuantity")!.Value);
            Assert.Null(discounts[1].Element(Ns + "ToQuantity"));
            Assert.Null(discounts[1].Element(Ns + "DiscountAmount"));
            Assert.Null(discounts[2].Element(Ns + "DiscountPercent"));
            Assert.Null(discounts[2].Element(Ns + "Quantity"));
            Assert.Equal("10", xml.Descendants(Ns + "PriceAmount").Single().Value);
        } finally { CultureInfo.CurrentCulture = before; }
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("negative-percent")]
    [InlineData("large-percent")]
    [InlineData("negative-amount")]
    [InlineData("large-amount")]
    [InlineData("zero")]
    [InlineData("minimum")]
    [InlineData("maximum")]
    [InlineData("reversed")]
    [InlineData("kind")]
    [InlineData("count")]
    public void InvalidDiscountsRejectWithoutChangingTheProject(string kind) {
        var discount = kind switch {
            "missing" => Discount() with { Percent = null, Amount = null },
            "negative-percent" => Discount() with { Percent = -1 },
            "large-percent" => Discount() with { Percent = 101 },
            "negative-amount" => Discount() with { Amount = -1 },
            "large-amount" => Discount() with { Amount = 11 },
            "zero" => Discount() with { Percent = 0 },
            "minimum" => Discount() with { MinimumQuantity = 0 },
            "maximum" => Discount() with { MaximumQuantity = 5 },
            "reversed" => Discount() with { MinimumQuantity = 10, MaximumQuantity = 5 },
            "kind" => Discount() with { Kind = (BookOnixDiscountKind)99 },
            _ => Discount()
        };
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        var options = Options(kind == "count" ? Enumerable.Repeat(discount, 17).ToArray() : [discount]);
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(options, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}

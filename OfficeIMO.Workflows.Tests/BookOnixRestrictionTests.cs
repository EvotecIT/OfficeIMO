using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixRestrictionTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixSalesRestriction Restriction() => new(BookOnixSalesRestrictionKind.RetailerExclusive) {
        Outlets = [new() { Name = "Example & Books", NameLanguageCode = "eng", Identifiers = [
            new(BookOnixSalesOutletScheme.Proprietary, "store-1") { SchemeName = "Publisher outlets" },
            new(BookOnixSalesOutletScheme.Onix, "AMZ"), new(BookOnixSalesOutletScheme.Gln, "1234567890128"),
            new(BookOnixSalesOutletScheme.San, "1234567")] }],
        Notes = [new("Only <this> outlet", "eng"), new("Wyłączność", "pol")],
        ValidFrom = new DateOnly(2026, 10, 1), ValidUntil = new DateOnly(2026, 10, 31)
    };
    private static BookOnixExportResult Export(params BookOnixSalesRestriction[] restrictions) => BookOnixTests.Project().ExportOnix(
        BookOnixTests.Options() with { Commercial = new() {
            SalesRights = [new(BookOnixSalesRightsKind.Exclusive, new() { Worldwide = true }) { Restrictions = restrictions }],
            Supplies = [new() { MarketReference = "world", Territory = new() { Worldwide = true }, SupplierName = "Supplier",
                SupplierRole = BookOnixSupplierRole.PublisherToCustomers, Availability = BookOnixAvailability.Available,
                Unpriced = BookOnixUnpricedKind.Free, Restrictions = restrictions }]
        } }, BookOnixTests.TestSchema());

    [Fact]
    public void RightsAndMarketRestrictionsUseTheSameStructureAndSurviveSelectiveUpdates() {
        var source = Export(Restriction()); var xml = XDocument.Load(new MemoryStream(source.Bytes));
        var restrictions = xml.Descendants(Ns + "SalesRestriction").ToArray();
        Assert.Equal(2, restrictions.Length); Assert.True(XNode.DeepEquals(restrictions[0], restrictions[1]));
        Assert.Equal(new[] { "01", "03", "04", "05" }, restrictions[0].Descendants(Ns + "SalesOutletIDType").Select(e => e.Value));
        Assert.Equal("Example & Books", restrictions[0].Descendants(Ns + "SalesOutletName").Single().Value);
        Assert.Equal(new[] { "Only <this> outlet", "Wyłączność" }, restrictions[0].Elements(Ns + "SalesRestrictionNote").Select(e => e.Value));
        Assert.Equal("20261001", restrictions[0].Element(Ns + "StartDate")!.Value);
        Assert.Equal("20261031", restrictions[0].Element(Ns + "EndDate")!.Value);
        var message = BookOnixMessage.CreateBlockUpdates([new(source) { ReplaceMarketReferences = ["world"] }], BookOnixTests.TestSchema());
        Assert.True(XNode.DeepEquals(restrictions[1], XDocument.Load(new MemoryStream(message.Bytes)).Descendants(Ns + "SalesRestriction").Single()));
    }

    [Theory]
    [InlineData(BookOnixSalesRestrictionKind.RetailerExclusiveOrOwnBrand)]
    [InlineData(BookOnixSalesRestrictionKind.RetailerExclusive)]
    [InlineData(BookOnixSalesRestrictionKind.RetailerOwnBrand)]
    [InlineData(BookOnixSalesRestrictionKind.RetailerException)]
    [InlineData(BookOnixSalesRestrictionKind.SelectedSubscriptionServices)]
    [InlineData(BookOnixSalesRestrictionKind.SubscriptionServiceExclusive)]
    public void DesignatedOutletRestrictionsRequireAnOutlet(BookOnixSalesRestrictionKind kind) =>
        Assert.Throws<ArgumentException>(() => Export(new BookOnixSalesRestriction(kind)));

    [Theory]
    [InlineData("unspecified")]
    [InlineData("unknown")]
    [InlineData("internal")]
    [InlineData("pod")]
    [InlineData("empty-outlet")]
    [InlineData("no-restrictions-outlet")]
    [InlineData("reversed")]
    [InlineData("untranslated")]
    [InlineData("duplicate-language")]
    [InlineData("long-note")]
    [InlineData("outlet-language")]
    public void UnsupportedOrIncompleteAssertionsFail(string kind) {
        var value = Restriction();
        value = kind switch {
            "unspecified" => new(BookOnixSalesRestrictionKind.Unspecified),
            "unknown" => new((BookOnixSalesRestrictionKind)999),
            "internal" => new((BookOnixSalesRestrictionKind)3),
            "pod" => new((BookOnixSalesRestrictionKind)17),
            "empty-outlet" => value with { Outlets = [new()] },
            "no-restrictions-outlet" => value with { Kind = BookOnixSalesRestrictionKind.NoRestrictions },
            "reversed" => value with { ValidUntil = new DateOnly(2026, 9, 30) },
            "untranslated" => value with { Notes = [new("one"), new("two", "pol")] },
            "duplicate-language" => value with { Notes = [new("one", "eng"), new("two", "eng")] },
            "long-note" => value with { Notes = [new(new string('x', 301))] },
            _ => value with { Outlets = [new() { NameLanguageCode = "eng", Identifiers = [new(BookOnixSalesOutletScheme.Onix, "AMZ")] }] }
        };
        Assert.ThrowsAny<ArgumentException>(() => Export(value));
    }

    [Theory]
    [InlineData(BookOnixSalesOutletScheme.Proprietary, "id", null)]
    [InlineData(BookOnixSalesOutletScheme.Onix, "amz", null)]
    [InlineData(BookOnixSalesOutletScheme.Onix, "AMZ", "unexpected")]
    [InlineData(BookOnixSalesOutletScheme.Gln, "123456789012", null)]
    [InlineData(BookOnixSalesOutletScheme.San, "123-4567", null)]
    public void IdentifierSchemesEnforceTheirDeclaredLexicalContract(BookOnixSalesOutletScheme scheme, string value, string? name) =>
        Assert.ThrowsAny<ArgumentException>(() => Export(Restriction() with { Outlets = [new() { Identifiers = [new(scheme, value) { SchemeName = name }] }] }));

    [Theory]
    [InlineData(BookOnixSalesRestrictionKind.LibrariesOnly, BookOnixSalesRestrictionKind.ExceptLibraries)]
    [InlineData(BookOnixSalesRestrictionKind.SchoolsOnly, BookOnixSalesRestrictionKind.ExceptSchools)]
    [InlineData(BookOnixSalesRestrictionKind.SubscriptionServicesOnly, BookOnixSalesRestrictionKind.ExceptSubscriptionServices)]
    [InlineData(BookOnixSalesRestrictionKind.OnlineRetailOnly, BookOnixSalesRestrictionKind.ExceptOnlineRetail)]
    [InlineData(BookOnixSalesRestrictionKind.EducationOnly, BookOnixSalesRestrictionKind.ExceptEducation)]
    [InlineData(BookOnixSalesRestrictionKind.NoRestrictions, BookOnixSalesRestrictionKind.LibrariesOnly)]
    public void OpposingAssertionsRequireNonOverlappingPeriods(BookOnixSalesRestrictionKind first, BookOnixSalesRestrictionKind second) {
        Assert.Throws<ArgumentException>(() => Export(new(first), new(second)));
        Export(new(first) { ValidUntil = new DateOnly(2026, 10, 1) }, new(second) { ValidFrom = new DateOnly(2026, 10, 2) });
        Assert.Throws<ArgumentException>(() => Export(new(first) { ValidUntil = new DateOnly(2026, 10, 1) }, new(second) { ValidFrom = new DateOnly(2026, 10, 1) }));
    }

    [Fact]
    public void CountAndAggregateTextBoundsPreventUnboundedRestrictionTrees() {
        Assert.Throws<ArgumentException>(() => Export(Enumerable.Repeat(new BookOnixSalesRestriction(BookOnixSalesRestrictionKind.LibrariesOnly), 33).ToArray()));
        Assert.Throws<ArgumentException>(() => Export(Restriction() with { Outlets = Enumerable.Repeat(new BookOnixSalesOutlet { Name = "outlet" }, 17).ToArray() }));
        Assert.Throws<ArgumentException>(() => Export(Restriction() with { Outlets = [new() { Identifiers = Enumerable.Range(0, 9).Select(i => new BookOnixSalesOutletIdentifier(BookOnixSalesOutletScheme.Proprietary, "id") { SchemeName = "scheme" + i }).ToArray() }] }));
        var rich = Restriction() with { Outlets = Enumerable.Range(0, 16).Select(i => new BookOnixSalesOutlet { Name = new string('n', 200),
            Identifiers = Enumerable.Range(0, 8).Select(j => new BookOnixSalesOutletIdentifier(BookOnixSalesOutletScheme.Proprietary, new string('v', 100)) { SchemeName = new string('s', 99) + j }).ToArray() }).ToArray() };
        var error = Assert.Throws<ArgumentException>(() => Export(Enumerable.Repeat(rich, 32).ToArray()));
        Assert.Contains("524288", error.Message);
    }
}

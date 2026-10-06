using System.Xml;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixMarketTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixSupply Supply(string? reference, string country) => new() {
        MarketReference = reference, Territory = new() { Countries = [country] }, SupplierName = "Supplier",
        SupplierRole = BookOnixSupplierRole.PublisherToCustomers, Availability = BookOnixAvailability.Available,
        Unpriced = BookOnixUnpricedKind.Free
    };
    private static BookOnixExportResult Export(params BookOnixSupply[] supplies) => BookOnixTests.Project().ExportOnix(
        BookOnixTests.Options() with { Commercial = new() {
            SalesRights = [new(BookOnixSalesRightsKind.Exclusive, new() { Worldwide = true })], Supplies = supplies
        } }, BookOnixTests.TestSchema());
    private static XElement Product(byte[] bytes) => XDocument.Load(new MemoryStream(bytes)).Root!.Element(Ns + "Product")!;
    private static BookOnixMessage Update(BookOnixBlockUpdate update) => BookOnixMessage.CreateBlockUpdates([update], BookOnixTests.TestSchema());

    [Fact]
    public void SelectiveReplacementPreservesWholeMarketAndOmissionDoesNotBecomeRemoval() {
        var source = Export(Supply("market.日本", "JP"), Supply("gb", "GB"));
        var before = source.Bytes.ToArray(); var selections = new[] { "gb", "market.日本" };
        var message = Update(new(source) { ReplaceMarketReferences = selections, ClearBlocks = [BookOnixBlock.CollateralDetail] });
        selections[0] = "changed";
        var supplies = Product(message.Bytes).Elements(Ns + "ProductSupply").ToArray();
        Assert.Equal(new[] { "gb", "market.日本" }, supplies.Select(e => e.Element(Ns + "MarketReference")!.Value));
        var originals = Product(source.Bytes).Elements(Ns + "ProductSupply").ToArray();
        Assert.True(XNode.DeepEquals(originals[1], supplies[0])); Assert.True(XNode.DeepEquals(originals[0], supplies[1]));
        Assert.Equal(before, source.Bytes); Assert.Same(source, Assert.Single(message.Products));
        var selected = Product(Update(new(source) { ReplaceMarketReferences = ["gb"] }).Bytes);
        Assert.Single(selected.Elements(Ns + "ProductSupply"));
        Assert.DoesNotContain(selected.Elements(Ns + "ProductSupply"), e => !e.Elements(Ns + "SupplyDetail").Any());
    }

    [Fact]
    public void ExplicitRemovalMayTargetAbsentPriorMarketAndKeepsOtherBlocksOrdered() {
        var source = Export(Supply("gb", "GB"));
        var xml = Product(Update(new(source) { ReplaceBlocks = [BookOnixBlock.PublishingDetail],
            ReplaceMarketReferences = ["gb"], RemoveMarketReferences = ["former & market"] }).Bytes);
        var supplies = xml.Elements(Ns + "ProductSupply").ToArray();
        Assert.Equal("former & market", Assert.Single(supplies[1].Elements()).Value);
        Assert.Equal(Ns + "MarketReference", Assert.Single(supplies[1].Elements()).Name);
        Assert.True(xml.Elements().ToList().IndexOf(xml.Element(Ns + "PublishingDetail")!) < xml.Elements().ToList().IndexOf(supplies[0]));
        Assert.Single(Product(Update(new(Export()) { RemoveMarketReferences = ["last-market"] }).Bytes).Elements(Ns + "ProductSupply"));
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData(null)]
    public void ExplicitInvalidSelectionReferencesFail(string? reference) => Assert.ThrowsAny<ArgumentException>(() =>
        Update(new(Export()) { RemoveMarketReferences = [reference!] }));

    [Fact]
    public void ReferenceLengthCharactersAndProductScopedUniquenessAreEnforced() {
        Export(Supply(new string('x', 100), "GB"));
        Assert.Throws<ArgumentException>(() => Export(Supply(new string('x', 101), "GB")));
        Assert.Throws<XmlException>(() => Export(Supply("bad\u0001", "GB")));
        Assert.Throws<ArgumentException>(() => Export(Supply("same", "GB"), Supply("same", "PL")));
        Assert.Throws<ArgumentException>(() => Update(new(Export()) { RemoveMarketReferences = [new string('x', 101)] }));
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("duplicate")]
    [InlineData("overlap")]
    [InlineData("whole")]
    [InlineData("unnamed")]
    [InlineData("mixed")]
    [InlineData("too-many")]
    [InlineData("null-list")]
    public void AmbiguousOrUnresolvableOperationsFail(string kind) {
        var source = kind == "unnamed" ? Export(Supply(null, "GB")) : kind == "mixed" ? Export(Supply("gb", "GB"), Supply(null, "PL")) : Export(Supply("gb", "GB"));
        var update = new BookOnixBlockUpdate(source) { ReplaceMarketReferences = ["gb"] };
        update = kind switch {
            "missing" => update with { ReplaceMarketReferences = ["GB"] },
            "duplicate" => update with { ReplaceMarketReferences = ["gb", "gb"] },
            "overlap" => update with { RemoveMarketReferences = ["gb"] },
            "whole" => update with { ReplaceBlocks = [BookOnixBlock.ProductSupply] },
            "too-many" => update with { RemoveMarketReferences = Enumerable.Range(0, 32).Select(i => "old" + i).ToArray() },
            "null-list" => update with { RemoveMarketReferences = null! },
            _ => update
        };
        Assert.ThrowsAny<ArgumentException>(() => Update(update));
    }

    [Fact]
    public void NamedMarketsCannotSilentlyAcquireAllMarketsReplacementSemantics() => Assert.Throws<ArgumentException>(() =>
        Update(new(Export(Supply("gb", "GB"))) { ReplaceBlocks = [BookOnixBlock.ProductSupply] }));
}

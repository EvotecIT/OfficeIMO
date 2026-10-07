using System.Xml;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixChangeTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixExportResult Export() => BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with {
        CollateralTexts = [new() { Type = BookOnixTextType.Description, Audiences = [BookOnixContentAudience.Unrestricted],
            Texts = [new("<p><strong>First</strong> <em>second</em></p><pre>  A\n B </pre>") { Format = BookOnixCollateralTextFormat.Xhtml }] }]
    }, BookOnixTests.TestSchema());
    private static XElement Product(byte[] bytes) => XDocument.Load(new MemoryStream(bytes)).Root!.Element(Ns + "Product")!;

    [Fact]
    public void ReplacementCopiesWholeBlockAndPreservesIdentityContentAndSourceEvidence() {
        var source = Export(); byte[] before = source.Bytes.ToArray();
        var selected = new[] { BookOnixBlock.CollateralDetail };
        var message = BookOnixMessage.CreateBlockUpdates([new(source) { ReplaceBlocks = selected }], BookOnixTests.TestSchema());
        selected[0] = BookOnixBlock.DescriptiveDetail;
        var product = Product(message.Bytes); var original = Product(source.Bytes);
        Assert.Equal(BookOnixMessageKind.BlockUpdates, message.Kind);
        Assert.Equal("04", product.Element(Ns + "NotificationType")!.Value);
        Assert.True(XNode.DeepEquals(original.Element(Ns + "CollateralDetail"), product.Element(Ns + "CollateralDetail")));
        Assert.True(XNode.DeepEquals(original.Element(Ns + "ProductIdentifier"), product.Element(Ns + "ProductIdentifier")));
        Assert.Equal(original.Element(Ns + "RecordReference")!.Value, product.Element(Ns + "RecordReference")!.Value);
        Assert.Null(product.Element(Ns + "DescriptiveDetail")); Assert.Null(product.Element(Ns + "PublishingDetail"));
        Assert.Same(source, Assert.Single(message.Products)); Assert.Equal(before, source.Bytes);
    }

    [Fact]
    public void ExplicitClearsAreEmptyAndOrderedWhileOmittedBlocksStayAbsent() {
        var message = BookOnixMessage.CreateBlockUpdates([new(Export()) { ClearBlocks = [
            BookOnixBlock.ProductionDetail, BookOnixBlock.RelatedMaterial, BookOnixBlock.ContentDetail,
            BookOnixBlock.PromotionDetail, BookOnixBlock.CollateralDetail] }], BookOnixTests.TestSchema());
        var blocks = Product(message.Bytes).Elements().Skip(3).ToArray();
        Assert.Equal(new[] { "CollateralDetail", "PromotionDetail", "ContentDetail", "RelatedMaterial", "ProductionDetail" }, blocks.Select(e => e.Name.LocalName));
        Assert.All(blocks, block => Assert.Empty(block.Nodes()));
    }

    [Theory]
    [InlineData("empty")]
    [InlineData("overlap")]
    [InlineData("duplicate")]
    [InlineData("undefined")]
    [InlineData("missing")]
    [InlineData("null")]
    public void AmbiguousOrMissingReplacementInstructionsFail(string kind) {
        var update = new BookOnixBlockUpdate(Export());
        update = kind switch {
            "overlap" => update with { ReplaceBlocks = [BookOnixBlock.CollateralDetail], ClearBlocks = [BookOnixBlock.CollateralDetail] },
            "duplicate" => update with { ReplaceBlocks = [BookOnixBlock.CollateralDetail, BookOnixBlock.CollateralDetail] },
            "undefined" => update with { ClearBlocks = [(BookOnixBlock)999] },
            "missing" => update with { ReplaceBlocks = [BookOnixBlock.RelatedMaterial] },
            "null" => update with { ReplaceBlocks = null! },
            _ => update
        };
        Assert.ThrowsAny<ArgumentException>(() => BookOnixMessage.CreateBlockUpdates([update], BookOnixTests.TestSchema()));
    }

    [Theory]
    [InlineData(BookOnixBlock.DescriptiveDetail)]
    [InlineData(BookOnixBlock.PublishingDetail)]
    [InlineData(BookOnixBlock.ProductSupply)]
    public void RequiredOrMarketBlocksCannotBeCleared(BookOnixBlock block) =>
        Assert.Throws<ArgumentException>(() => BookOnixMessage.CreateBlockUpdates([new(Export()) { ClearBlocks = [block] }], BookOnixTests.TestSchema()));

    [Fact]
    public void DeletionRetainsIdentityAndTranslatedReasonsButNoPublicationBlocks() {
        var source = Export();
        var message = BookOnixMessage.CreateDeletions([new(source) { Reasons = [new("Issued <in> error & duplicated", "eng"), new("Błędny rekord", "pol")] }], BookOnixTests.TestSchema());
        var product = Product(message.Bytes);
        Assert.Equal(BookOnixMessageKind.Deletions, message.Kind);
        Assert.Equal(new[] { "RecordReference", "NotificationType", "DeletionText", "DeletionText", "ProductIdentifier" }, product.Elements().Select(e => e.Name.LocalName));
        Assert.Equal("05", product.Element(Ns + "NotificationType")!.Value);
        Assert.Equal(new[] { "Issued <in> error & duplicated", "Błędny rekord" }, product.Elements(Ns + "DeletionText").Select(e => e.Value));
        Assert.Equal(new[] { "eng", "pol" }, product.Elements(Ns + "DeletionText").Select(e => (string?)e.Attribute("language")));
        Assert.Same(source, Assert.Single(message.Products));
        Assert.Empty(Product(BookOnixMessage.CreateDeletions([new(source)], BookOnixTests.TestSchema()).Bytes).Elements(Ns + "DeletionText"));
    }

    [Theory]
    [InlineData("blank")]
    [InlineData("long")]
    [InlineData("missing-language")]
    [InlineData("duplicate-language")]
    [InlineData("bad-language")]
    [InlineData("count")]
    public void InvalidDeletionReasonsFail(string kind) {
        BookOnixDeletionReason[] reasons = kind switch {
            "blank" => [new(" ")], "long" => [new(new string('x', 101))],
            "missing-language" => [new("one"), new("two", "pol")],
            "duplicate-language" => [new("one", "eng"), new("two", "eng")],
            "bad-language" => [new("one", "en")],
            _ => Enumerable.Repeat(new BookOnixDeletionReason("reason", "eng"), 17).ToArray()
        };
        Assert.ThrowsAny<ArgumentException>(() => BookOnixMessage.CreateDeletions([new(Export()) { Reasons = reasons }], BookOnixTests.TestSchema()));
    }

    [Fact]
    public void InvalidXmlReasonIsNotSilentlySanitized() => Assert.Throws<XmlException>(() =>
        BookOnixMessage.CreateDeletions([new(Export()) { Reasons = [new("bad\u0001text")] }], BookOnixTests.TestSchema()));

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ChangeFactoriesRetainIntegritySchemaIdentityAndCancellationGates(bool deletion) {
        var source = Export();
        BookOnixMessage Compose(System.Xml.Schema.XmlSchemaSet schema, CancellationToken token = default, bool duplicate = false) => deletion
            ? BookOnixMessage.CreateDeletions(duplicate ? [new(source), new(source)] : [new(source)], schema, token)
            : BookOnixMessage.CreateBlockUpdates(duplicate ? [new(source) { ReplaceBlocks = [BookOnixBlock.DescriptiveDetail] }, new(source) { ReplaceBlocks = [BookOnixBlock.DescriptiveDetail] }] : [new(source) { ReplaceBlocks = [BookOnixBlock.DescriptiveDetail] }], schema, token);
        Assert.Throws<ArgumentException>(() => Compose(new()));
        Assert.Throws<InvalidDataException>(() => Compose(BookOnixTests.TestSchema(true)));
        Assert.Throws<ArgumentException>(() => Compose(BookOnixTests.TestSchema(), duplicate: true));
        Assert.Throws<OperationCanceledException>(() => Compose(BookOnixTests.TestSchema(), new CancellationToken(true)));
        source.Publication.Bytes[0] ^= 1;
        Assert.Throws<InvalidDataException>(() => Compose(BookOnixTests.TestSchema()));
    }
}

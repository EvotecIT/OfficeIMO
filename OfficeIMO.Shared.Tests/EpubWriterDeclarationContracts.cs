using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Epub;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubWriterDeclarationContracts {
    [Theory]
    [InlineData("first", "text/css", false)]
    [InlineData("style", "text/css", false)]
    [InlineData("overlay", "image/png", false)]
    [InlineData("overlay", "application/smil+xml", true)]
    public async Task OverlayAssociations_ValidateFinalSourceAndTargetTypesBeforeSaving(string targetId, string targetType, bool changeSource) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.AddResource("overlay", "EPUB/overlay.smil", targetType, Encoding.UTF8.GetBytes("<smil xmlns='http://www.w3.org/ns/SMIL' version='3.0'><body/></smil>"));
        book.Manifest.Single(item => item.Id == (changeSource ? "style" : "first")).MediaOverlayId = targetId;
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<InvalidDataException>(() => book.Save(destination));
        await Assert.ThrowsAsync<InvalidDataException>(() => book.SaveAsync(destination));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
    }

    [Fact]
    public void OverlayAssociations_RetainSmilRelationshipsAndRevalidateMutatedTargets() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.AddResource("overlay", "EPUB/overlay.smil", "application/smil+xml", Encoding.UTF8.GetBytes("<smil xmlns='http://www.w3.org/ns/SMIL' version='3.0'><body/></smil>"));
        book.Manifest.Single(item => item.Id == "first").MediaOverlayId = "overlay";
        byte[] source = book.Write().Bytes;
        book = EpubPublication.Load(new MemoryStream(source));
        Assert.Equal(source, book.Write().Bytes);
        book.Title = "Edited";
        Assert.Equal("overlay", EpubPublication.Load(new MemoryStream(book.Write().Bytes)).Manifest.Single(item => item.Id == "first").MediaOverlayId);
        book.Manifest.Single(item => item.Id == "overlay").MediaType = "image/png";
        Assert.Throws<InvalidDataException>(() => book.Write());
    }

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void CoverSelection_RevalidatesFinalTypeAfterImportAndMutation(EpubVersion version) {
        EpubPublication book = EpubWritingContracts.CreateBook(version);
        book.AddResource("cover", "EPUB/cover.png", "image/png", new byte[] { 1 });
        book.SetCoverImage("cover");
        book = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        book.Manifest.Single(item => item.Id == "cover").MediaType = "text/css";
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<InvalidDataException>(() => book.Save(destination));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CoverProperties_RejectNonImagesAndMultipleSelections(bool multiple) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        if (multiple) {
            book.AddResource("cover", "EPUB/cover.png", "image/png", new byte[] { 1 }, "cover-image");
            book.AddResource("other", "EPUB/other.png", "image/png", new byte[] { 1 }, "cover-image");
        } else book.Manifest.Single(item => item.Id == "style").Properties = "cover-image";
        Assert.Throws<InvalidDataException>(() => book.Write());
    }

    [Theory]
    [InlineData("_", "https://example.test/")]
    [InlineData("custom", "http://idpf.org/epub/vocab/package/#")]
    [InlineData("custom", "http://idpf.org/epub/vocab/package/item/#")]
    [InlineData("custom", "http://idpf.org/epub/vocab/package/itemref/#")]
    [InlineData("custom", "http://idpf.org/epub/vocab/package/link/#")]
    [InlineData("custom", "http://idpf.org/epub/vocab/structure/#")]
    [InlineData("custom", "http://purl.org/dc/elements/1.1/")]
    [InlineData("custom", "https://example.test/vocabulary with spaces/")]
    public void VocabularyDeclarations_RejectProhibitedMappingsBeforeMutation(string prefix, string uri) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.DeclareVocabularyPrefix(prefix, uri));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void VocabularyDeclarations_AllowCustomUnderscoreNamesAndReservedVocabularyAliases() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.DeclareVocabularyPrefix("_custom", "https://example.test/vocabulary/");
        book.AddMetadataProperty("_custom:value", "Custom metadata");
        book.DeclareVocabularyPrefix("mySchema", "http://schema.org/");
        book.SetMetadataProperty("mySchema:accessMode", "textual");
        EpubPublication output = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Contains("_custom: https://example.test/vocabulary/", (string?)output.GetPackageXml().Root!.Attribute("prefix"));
    }

    [Theory]
    [InlineData("unknown:term")]
    [InlineData("schema:")]
    public void MetadataProperties_RejectUndeclaredOrEmptyPrefixedReferencesBeforeMutation(string property) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.SetMetadataProperty(property, "Value"));
        Assert.Throws<ArgumentException>(() => book.AddMetadataProperty(property, "Value"));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData(EpubVersion.Epub2, false)]
    [InlineData(EpubVersion.Epub2, true)]
    [InlineData(EpubVersion.Epub3, false)]
    public void SvgSpine_RequiresVersionAppropriateContentFallback(EpubVersion version, bool withFallback) {
        EpubPublication book = EpubWritingContracts.CreateBook(version);
        EpubManifestItem svg = book.AddResource("vector", "EPUB/vector.svg", "image/svg+xml",
            Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text>Vector</text></svg>"));
        if (withFallback) svg.FallbackId = "first";
        book.AddSpineItem("vector");
        if (version == EpubVersion.Epub2 && !withFallback) Assert.Throws<NotSupportedException>(() => book.Write());
        else {
            EpubPublication output = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
            Assert.Contains(output.Spine, item => item.ManifestId == "vector");
            Assert.Equal(withFallback ? "first" : null, output.Manifest.Single(item => item.Id == "vector").FallbackId);
        }
    }
}

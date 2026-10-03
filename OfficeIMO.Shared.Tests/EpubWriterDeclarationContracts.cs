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
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void CoverReselection_UpdatesAllRetainedLegacyDeclarationsAndReleasesOldImage(EpubVersion version) {
        EpubPublication book = EpubWritingContracts.CreateBook(version);
        book.AddResource("old-cover", "EPUB/old.png", "image/png", new byte[] { 1 });
        book.AddResource("new-cover", "EPUB/new.png", "image/png", new byte[] { 2 });
        book.SetCoverImage("old-cover");
        XDocument package = book.GetPackageXml();
        XNamespace opf = "http://www.idpf.org/2007/opf";
        XElement metadata = package.Root!.Element(opf + "metadata")!;
        if (version == EpubVersion.Epub3) metadata.Add(new XElement(opf + "meta", new XAttribute("name", "cover"), new XAttribute("content", "old-cover")));
        metadata.Add(new XElement(opf + "meta", new XAttribute("name", "cover"), new XAttribute("content", "old-cover"), new XAttribute("id", "compat-cover")));
        byte[] source = EpubWritingContracts.ReplaceEntry(book.Write().Bytes, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()));
        book = EpubPublication.Load(new MemoryStream(source));
        book.SetCoverImage("new-cover");
        book.RemoveResource("old-cover");
        EpubPublication output = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.All(output.GetPackageXml().Root!.Element(opf + "metadata")!.Elements(opf + "meta").Where(meta => (string?)meta.Attribute("name") == "cover"),
            meta => Assert.Equal("new-cover", (string?)meta.Attribute("content")));
        Assert.Contains(output.GetPackageXml().Descendants(opf + "meta"), meta => (string?)meta.Attribute("id") == "compat-cover");
        Assert.DoesNotContain(output.Manifest, item => item.Id == "old-cover");
    }

    [Theory]
    [InlineData("manifest", "unknown:term")]
    [InlineData("spine", "schema:")]
    [InlineData("resource", "unknown:term")]
    [InlineData("position", "schema:")]
    public void PropertyAuthoring_RejectsInvalidTokensBeforeMutation(string route, string properties) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        byte[] before = book.Write().Bytes;
        Action mutation = route == "manifest" ? () => book.Manifest.Single(item => item.Id == "first").Properties = properties :
            route == "spine" ? () => book.Spine[0].Properties = properties :
            route == "resource" ? () => book.AddResource("extra", "EPUB/extra.bin", "application/octet-stream", new byte[] { 1 }, properties) :
            () => book.AddSpineItem("style", properties: properties);
        Assert.Throws<ArgumentException>(mutation);
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData("_", "https://example.test/vocabulary/")]
    [InlineData("custom", "http://idpf.org/epub/vocab/package/#")]
    [InlineData("custom", "http://purl.org/dc/elements/1.1/")]
    public void ImportedProhibitedMappings_CannotAuthorizeNewMetadataOrPropertyTokens(string prefix, string uri) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        XDocument package = book.GetPackageXml();
        package.Root!.SetAttributeValue("prefix", prefix + ": " + uri);
        byte[] source = EpubWritingContracts.ReplaceEntry(book.Write().Bytes, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()));
        book = EpubPublication.Load(new MemoryStream(source));
        Assert.Equal(source, book.Write().Bytes);
        Assert.Throws<ArgumentException>(() => book.SetMetadataProperty(prefix + ":term", "Value"));
        Assert.Throws<ArgumentException>(() => book.AddMetadataProperty(prefix + ":term", "Value"));
        Assert.Throws<ArgumentException>(() => book.Manifest.Single(item => item.Id == "first").Properties = prefix + ":term");
        Assert.Equal(source, book.Write().Bytes);
    }

    [Fact]
    public void PropertyAuthoring_RetainsImportedUnknownTokensAndAcceptsDeclaredCustomTokens() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        XDocument package = book.GetPackageXml();
        XNamespace opf = "http://www.idpf.org/2007/opf";
        package.Root!.Element(opf + "manifest")!.Elements(opf + "item").Single(item => (string?)item.Attribute("id") == "first").SetAttributeValue("properties", "legacy:term");
        byte[] source = EpubWritingContracts.ReplaceEntry(book.Write().Bytes, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()));
        book = EpubPublication.Load(new MemoryStream(source));
        book.Title = "Edited with retained extensions";
        book.DeclareVocabularyPrefix("custom", "https://example.test/vocabulary/");
        book.Manifest.Single(item => item.Id == "style").Properties = "custom:term";
        book.Spine[0].Properties = "custom:position";
        EpubPublication output = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal("legacy:term", output.Manifest.Single(item => item.Id == "first").Properties);
        Assert.Equal("custom:term", output.Manifest.Single(item => item.Id == "style").Properties);
        Assert.Equal("custom:position", output.Spine[0].Properties);
    }

    [Theory]
    [InlineData("image/png", false)]
    [InlineData("text/css", false)]
    [InlineData("application/pdf", true)]
    public void Epub2Spine_SeparatesEmbeddedCoreResourcesFromForeignContentFallbacks(string mediaType, bool accepted) {
        EpubPublication book = EpubWritingContracts.CreateBook(EpubVersion.Epub2);
        EpubManifestItem item = book.AddResource("foreign", "EPUB/foreign.bin", mediaType, new byte[] { 1 });
        item.FallbackId = "first";
        book.AddSpineItem("foreign", linear: false);
        if (accepted) Assert.Contains(EpubPublication.Load(new MemoryStream(book.Write().Bytes)).Spine, position => position.ManifestId == "foreign");
        else Assert.Throws<NotSupportedException>(() => book.Write());
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
        if (version == EpubVersion.Epub2) Assert.Throws<NotSupportedException>(() => book.Write());
        else {
            EpubPublication output = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
            Assert.Contains(output.Spine, item => item.ManifestId == "vector");
            Assert.Equal(withFallback ? "first" : null, output.Manifest.Single(item => item.Id == "vector").FallbackId);
        }
    }
}

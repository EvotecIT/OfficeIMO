using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Fact]
    public void JumbfStoreRejectsMalformedTrailingChildBox() {
        byte[] manifest = CreateManifestStore();
        Array.Resize(ref manifest, manifest.Length + 4);
        WriteBigEndian(manifest, 0, manifest.Length);
        byte[] png = CreatePngWithManifest(manifest);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void JumbfStoreRequiresAtLeastOneManifestSuperbox() {
        byte[] storeUuid = { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] description = CreateBox("jumd", Join(storeUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa\0")));
        byte[] png = CreatePngWithManifest(CreateBox("jumb", description));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void JumbfManifestSuperboxRequiresAClaimBox() {
        byte[] storeUuid = { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] manifestUuid = { 0x63, 0x32, 0x6D, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] storeDescription = CreateBox("jumd", Join(storeUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa\0")));
        byte[] manifestDescription = CreateBox("jumd", Join(manifestUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("manifest\0")));
        byte[] descriptionOnlyManifest = CreateBox("jumb", manifestDescription);
        byte[] png = CreatePngWithManifest(CreateBox("jumb", Join(storeDescription, descriptionOnlyManifest)));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void PngChunkCountIsBoundedBeforeCarrierProcessing() {
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 }),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            png,
            "fixture.png",
            new OfficeProvenanceOptions { MaxContainerEntries = 1 }));
    }

    [Fact]
    public void WebpChunkCountIsBoundedBeforeCarrierProcessing() {
        byte[] webp = CreateWebp(
            CreateRiffChunk("VP8 ", Array.Empty<byte>()),
            CreateRiffChunk("VP8 ", Array.Empty<byte>()),
            CreateRiffChunk("VP8 ", Array.Empty<byte>()));
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxContainerEntries = 2;

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            webp,
            "fixture.webp",
            new OfficeProvenanceOptions { MaxContainerEntries = 2 }));
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceRemover.Remove(
            webp,
            "fixture.webp",
            removalOptions));
    }

    [Fact]
    public void GifBlocksAndSubBlocksShareTheContainerEntryLimit() {
        byte[] gif = Join(
            Encoding.ASCII.GetBytes("GIF89a"),
            new byte[] { 1, 0, 1, 0, 0, 0, 0 },
            new byte[] { 0x21, 0xFE, 1, (byte)'a', 0 },
            new byte[] { 0x21, 0xFE, 1, (byte)'b', 0 },
            new byte[] { 0x3B });

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            gif,
            "fixture.gif",
            new OfficeProvenanceOptions { MaxContainerEntries = 6 }));
        Assert.Empty(OfficeProvenanceInspector.Inspect(
            gif,
            "fixture.gif",
            new OfficeProvenanceOptions { MaxContainerEntries = 7 }).Evidence);
    }

    [Fact]
    public void JpegMarkerCountIsBoundedBeforeMetadataProcessing() {
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE0, Array.Empty<byte>()),
            CreateJpegSegment(0xE0, Array.Empty<byte>()),
            new byte[] { 0xFF, 0xD9 });

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            jpeg,
            "fixture.jpg",
            new OfficeProvenanceOptions { MaxContainerEntries = 2 }));
        Assert.Empty(OfficeProvenanceInspector.Inspect(
            jpeg,
            "fixture.jpg",
            new OfficeProvenanceOptions { MaxContainerEntries = 4 }).Evidence);
    }

    [Fact]
    public void ConcatenatedJpegImagesShareTheContainerEntryLimit() {
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 },
            new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 },
            new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 });
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxContainerEntries = 2;

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            jpeg,
            "fixture.jpg",
            new OfficeProvenanceOptions { MaxContainerEntries = 2 }));
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceRemover.Remove(
            jpeg,
            "fixture.jpg",
            removalOptions));
    }

    [Fact]
    public void ConcatenatedJpegLookaheadDoesNotDoubleCountMarkers() {
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 },
            new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(
            jpeg,
            "fixture.jpg",
            new OfficeProvenanceOptions { MaxContainerEntries = 4 });

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void FragmentedJpegAcceptsAnExtendedSizeStoreDescriptionBox() {
        byte[] storeUuid = { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] manifestUuid = { 0x63, 0x32, 0x6D, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] claimUuid = { 0x63, 0x32, 0x63, 0x6C, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] storeDescription = CreateExtendedBox("jumd", Join(storeUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa\0")));
        byte[] manifestDescription = CreateBox("jumd", Join(manifestUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("m\0")));
        byte[] assertionStoreUuid = { 0x63, 0x32, 0x61, 0x73, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] assertionStoreDescription = CreateBox("jumd", Join(assertionStoreUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.assertions\0")));
        byte[] assertionUuid = { 0x63, 0x32, 0x61, 0x63, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] assertionDescription = CreateBox("jumd", Join(assertionUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.test\0")));
        byte[] assertionStore = CreateBox("jumb", Join(assertionStoreDescription,
            CreateBox("jumb", Join(assertionDescription, CreateBox("cbor", new byte[] { 0xA0 })))));
        byte[] claimDescription = CreateBox("jumd", Join(claimUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.claim\0")));
        byte[] claim = CreateBox("jumb", Join(claimDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] signatureUuid = { 0x63, 0x32, 0x63, 0x73, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] signatureDescription = CreateBox("jumd", Join(signatureUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.signature\0")));
        byte[] signature = CreateBox("jumb", Join(signatureDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] manifest = CreateBox("jumb", Join(storeDescription, CreateBox("jumb", Join(manifestDescription, assertionStore, claim, signature))));
        byte[] jpeg = CreateValidJpeg(
            CreateJpegApp11(manifest, 0, 46, instance: 11, sequence: 1),
            CreateJpegApp11(manifest, 46, manifest.Length - 46, instance: 11, sequence: 2));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.True(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void JumbfChildBoxesShareTheContainerEntryLimit() {
        byte[] png = CreatePngWithManifest(CreateManifestStore());

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            png,
            "fixture.png",
            new OfficeProvenanceOptions { MaxContainerEntries = 14 }));
        OfficeProvenanceReport accepted = OfficeProvenanceInspector.Inspect(
            png,
            "fixture.png",
            new OfficeProvenanceOptions { MaxContainerEntries = 15 });

        Assert.True(Assert.Single(accepted.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void ManifestWithoutClaimSignatureIsStructurallyInvalid() {
        byte[] png = CreatePngWithManifest(CreateManifestStore(includeSignature: false));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void JumbfStoreRejectsAnUnrecognizedTrailingChildBox() {
        byte[] manifest = CreateManifestStore();
        byte[] unrecognized = CreateBox("free", new byte[] { 1 });
        manifest = Join(manifest, unrecognized);
        WriteBigEndian(manifest, 0, manifest.Length);
        byte[] png = CreatePngWithManifest(manifest);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void VariationSelectorWrapperWinsOverGenericTextExtension() {
        byte[] wrapper = CreateTextWrapper(CreateManifestStore());

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(wrapper, "fixture.txt");

        Assert.Equal(OfficeProvenanceAssetFormat.UnstructuredText, report.Format);
        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void GenericRemovalPreservesHtmlTextThatLooksLikeAStructuredCarrier() {
        byte[] html = Encoding.UTF8.GetBytes(
            "<!doctype html><html><body><pre>\n-----BEGIN C2PA MANIFEST-----\nhttps://example.com/manifest.c2pa\n-----END C2PA MANIFEST-----\n</pre></body></html>");

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(html, "fixture.html");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(html, "fixture.html");

        Assert.Equal(OfficeProvenanceAssetFormat.Html, report.Format);
        Assert.Empty(report.Evidence);
        Assert.Equal(html, result.ToArray());
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void StructuredExternalUriRequiresStrictUtf8() {
        byte[] text = Join(
            Encoding.ASCII.GetBytes("-----BEGIN C2PA MANIFEST-----\nhttps://example.com/"),
            new byte[] { 0xFF },
            Encoding.ASCII.GetBytes("\n-----END C2PA MANIFEST-----\n"));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(text, "fixture.md");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(text, "fixture.md");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.Equal(text, result.ToArray());
    }

    [Fact]
    public void StructuredTextDelimiterCandidatesShareTheContainerEntryLimit() {
        byte[] text = Encoding.ASCII.GetBytes(
            "-----BEGIN C2PA MANIFEST-----\n" +
            "-----BEGIN C2PA MANIFEST-----\n" +
            "-----BEGIN C2PA MANIFEST-----\n" +
            "https://example.test/manifest.c2pa\n" +
            "-----END C2PA MANIFEST-----\n");
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxContainerEntries = 2;

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            text,
            "fixture.md",
            new OfficeProvenanceOptions { MaxContainerEntries = 2 }));
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceRemover.Remove(
            text,
            "fixture.md",
            removalOptions));
    }

    [Fact]
    public void XmpDigitalSourceTypeRequiresAnRdfDescriptionContext() {
        byte[] packet = Encoding.UTF8.GetBytes(
            "<root xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\" iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), packet)),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.Empty(report.Evidence);
        Assert.Equal(jpeg, result.ToArray());
    }

    [Fact]
    public void MalformedOversizedTextWrapperIsRemovedAsOneCompleteRun() {
        byte[] wrapper = CreateTextWrapper(CreateManifestStore());
        byte[] suffix = Encoding.UTF8.GetBytes("tail");
        byte[] text = Join(wrapper, suffix);
        var options = new OfficeProvenanceRemovalOptions { RequireStructurallyValidCarrier = false };
        options.Limits.MaxManifestBytes = 1;

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(text, "fixture.txt", options);

        Assert.Equal(suffix, result.ToArray());
        Assert.Single(result.Changes);
    }

    [Fact]
    public void VariationSelectorWrappersHonorTheContainerEntryLimit() {
        byte[] wrapper = CreateTextWrapper(CreateManifestStore());
        var inspectOptions = new OfficeProvenanceOptions { MaxContainerEntries = 128 };
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxContainerEntries = 128;

        Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceInspector.Inspect(wrapper, "fixture.txt", inspectOptions));
        Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceRemover.Remove(wrapper, "fixture.txt", removalOptions));
    }

    [Fact]
    public void JpegInspectsAndRemovesAiDeclarationFromAdobeExtendedXmp() {
        byte[] extendedPacket = Encoding.UTF8.GetBytes(
            "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><rdf:Description " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></rdf:RDF>");
        string guid = ComputeMd5(extendedPacket);
        byte[] standardPacket = Encoding.UTF8.GetBytes(
            $"<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:xmpNote=\"http://ns.adobe.com/xmp/note/\" " +
            $"xmpNote:HasExtendedXMP=\"{guid}\"/>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), standardPacket)),
            CreateExtendedXmpSegment(guid, extendedPacket),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceReport before = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");
        OfficeProvenanceReport after = OfficeProvenanceInspector.Inspect(result.ToArray(), "fixture.jpg");

        Assert.Single(before.Evidence);
        Assert.True(result.WasChanged);
        Assert.Empty(after.Evidence);
    }

    [Fact]
    public void JpegRejectsExtendedXmpWhosePacketDoesNotMatchTheReferencedDigest() {
        byte[] originalPacket = Encoding.UTF8.GetBytes(
            "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\"/>");
        string guid = ComputeMd5(originalPacket);
        byte[] substitutedPacket = Encoding.UTF8.GetBytes(
            "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><rdf:Description iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></rdf:RDF>");
        byte[] standardPacket = Encoding.UTF8.GetBytes(
            $"<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:xmpNote=\"http://ns.adobe.com/xmp/note/\" xmpNote:HasExtendedXMP=\"{guid}\"/>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), standardPacket)),
            CreateExtendedXmpSegment(guid, substitutedPacket),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.Empty(report.Evidence);
        Assert.Contains(report.Diagnostics, item => item.Contains("digest", StringComparison.OrdinalIgnoreCase));
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void XmpNodeBudgetAppliesBeforeStandardPacketMaterialization() {
        byte[] packet = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><rdf:RDF><rdf:Description>" +
            string.Concat(Enumerable.Repeat("<x:n/>", 16)) +
            "<iptc:DigitalSourceType>trainedAlgorithmicMedia</iptc:DigitalSourceType></rdf:Description></rdf:RDF></x:xmpmeta>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), packet)),
            new byte[] { 0xFF, 0xD9 });

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            jpeg, "fixture.jpg", new OfficeProvenanceOptions { MaxContainerEntries = 8 }));
        OfficeProvenanceReport accepted = OfficeProvenanceInspector.Inspect(
            jpeg, "fixture.jpg", new OfficeProvenanceOptions { MaxContainerEntries = 64 });

        Assert.Single(accepted.Evidence);
    }

    [Fact]
    public void SvgContentWinsOverGenericXmlFileExtension() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:x=\"adobe:ns:meta/\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><metadata><x:xmpmeta><rdf:RDF><rdf:Description iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></rdf:RDF></x:xmpmeta></metadata></svg>");

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(svg, "fixture.xml");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.xml");

        Assert.Equal(OfficeProvenanceAssetFormat.Svg, report.Format);
        Assert.Single(report.Evidence);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void LeadingHtmlCommentsDoNotExposeStructuredTextCarriers() {
        string text = "<!-- legal notice -->\n<!doctype html><html><body><pre>-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) +
            "\n-----END C2PA MANIFEST-----</pre></body></html>";
        byte[] data = Encoding.UTF8.GetBytes(text);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(data, "fixture.txt");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(data, "fixture.txt");

        Assert.Equal(OfficeProvenanceAssetFormat.Html, report.Format);
        Assert.False(result.WasChanged);
        Assert.Equal(data, result.ToArray());
    }

    [Fact]
    public void SvgIgnoresXmpMarkupOutsideSvgMetadata() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:x=\"adobe:ns:meta/\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" +
            "<foreignObject><x:xmpmeta><rdf:RDF><rdf:Description iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></rdf:RDF></x:xmpmeta></foreignObject></svg>");

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(svg, "fixture.svg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
        Assert.Equal(svg, result.ToArray());
    }

    [Fact]
    public void Zip64EntryCountIsRejectedBeforeDirectoryMaterialization() {
        byte[] package = CreateZip64CountOnlyPackage(5000);

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            package,
            "fixture.zip",
            new OfficeProvenanceOptions { MaxContainerEntries = 10 }));
    }

    [Fact]
    public void ClassicZipMayContainExactlyTheSentinelEntryCountWithoutZip64Metadata() {
        byte[] endOfDirectory = new byte[22];
        WriteLittleEndian(endOfDirectory, 0, 0x06054B50U);
        WriteLittleEndian16(endOfDirectory, 8, ushort.MaxValue);
        WriteLittleEndian16(endOfDirectory, 10, ushort.MaxValue);

        OfficeProvenanceZip.ValidateEntryCount(endOfDirectory, ushort.MaxValue);
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceZip.ValidateEntryCount(endOfDirectory, ushort.MaxValue - 1));
    }

    [Fact]
    public void ZipRewritePreservesExternalAttributes() {
        byte[] localExtraField = { 0xFE, 0xCA, 0x03, 0x00, 0x10, 0x20, 0x30 };
        byte[] centralExtraField = { 0xFE, 0xCA, 0x02, 0x00, 0x40, 0x50 };
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                ZipArchiveEntry manifest = archive.CreateEntry("META-INF/content_credential.c2pa");
                using (Stream target = manifest.Open()) WriteAll(target, CreateManifestStore());
                ZipArchiveEntry script = archive.CreateEntry("bin/run.sh");
                script.ExternalAttributes = unchecked((int)0x81ED0000);
                using (Stream target = script.Open()) WriteAll(target, Encoding.UTF8.GetBytes("#!/bin/sh\n"));
            }
            package = AddEntryExtraFields(stream.ToArray(), "bin/run.sh", localExtraField, centralExtraField);
            package = AddCentralDirectoryComment(package, "bin/run.sh", Encoding.UTF8.GetBytes("keep-comment"));
            package = AddArchiveComment(package, Encoding.UTF8.GetBytes("keep-archive-comment"));
            int sourceCentralHeader = FindSignature(package, 0x02014B50u, "bin/run.sh");
            WriteLittleEndian16(package, sourceCentralHeader + 36, 1);
        }

        Assert.Equal("keep-archive-comment", Encoding.UTF8.GetString(ReadArchiveComment(package)));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip");
        using var rewritten = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);

        Assert.Equal(unchecked((int)0x81ED0000), rewritten.GetEntry("bin/run.sh")!.ExternalAttributes);
        int centralHeader = FindSignature(result.ToArray(), 0x02014B50u, "bin/run.sh");
        Assert.Equal(3, result.ToArray()[centralHeader + 5]);
        Assert.Equal(localExtraField, ReadLocalExtraField(result.ToArray(), centralHeader));
        Assert.Equal(centralExtraField, ReadCentralExtraField(result.ToArray(), centralHeader));
        Assert.Equal("keep-comment", Encoding.UTF8.GetString(ReadCentralDirectoryComment(result.ToArray(), centralHeader)));
        Assert.Equal("keep-archive-comment", Encoding.UTF8.GetString(ReadArchiveComment(result.ToArray())));
        Assert.Equal(1, BitConverter.ToUInt16(result.ToArray(), centralHeader + 36));
    }

    [Fact]
    public void ZipRewriteTranscodesLegacyEntryCommentsToUtf8() {
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                using (Stream manifest = archive.CreateEntry("META-INF/content_credential.c2pa").Open()) WriteAll(manifest, CreateManifestStore());
                using (Stream keep = archive.CreateEntry("keep.txt").Open()) WriteAll(keep, Encoding.UTF8.GetBytes("keep"));
            }
            package = AddCentralDirectoryComment(stream.ToArray(), "keep.txt", new byte[] { 0x82 });
        }
        int sourceCentralHeader = FindSignature(package, 0x02014B50u, "keep.txt");
        WriteLittleEndian16(package, sourceCentralHeader + 8,
            (ushort)(BitConverter.ToUInt16(package, sourceCentralHeader + 8) & ~0x0800));
        uint localHeaderOffset = BitConverter.ToUInt32(package, sourceCentralHeader + 42);
        WriteLittleEndian16(package, checked((int)localHeaderOffset) + 6,
            (ushort)(BitConverter.ToUInt16(package, checked((int)localHeaderOffset) + 6) & ~0x0800));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip");
        int centralHeader = FindSignature(result.ToArray(), 0x02014B50u, "keep.txt");

        Assert.NotEqual(0, BitConverter.ToUInt16(result.ToArray(), centralHeader + 8) & 0x0800);
        Assert.Equal("é", Encoding.UTF8.GetString(ReadCentralDirectoryComment(result.ToArray(), centralHeader)));
    }

    [Fact]
    public void ZipRewriteDropsFilenameDependentUnicodePathExtras() {
        byte[] retainedExtraField = { 0xFE, 0xCA, 0x01, 0x00, 0x42 };
        byte[] unicodeName = Encoding.UTF8.GetBytes("renamed.txt");
        byte[] unicodePathExtraField = new byte[9 + unicodeName.Length];
        WriteLittleEndian16(unicodePathExtraField, 0, 0x7075);
        WriteLittleEndian16(unicodePathExtraField, 2, checked((ushort)(5 + unicodeName.Length)));
        unicodePathExtraField[4] = 1;
        WriteLittleEndian(unicodePathExtraField, 5, 0xDEADBEEFu);
        Buffer.BlockCopy(unicodeName, 0, unicodePathExtraField, 9, unicodeName.Length);
        byte[] sourceExtraField = Join(retainedExtraField, unicodePathExtraField);
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                using (Stream manifest = archive.CreateEntry("META-INF/content_credential.c2pa").Open()) WriteAll(manifest, CreateManifestStore());
                using (Stream keep = archive.CreateEntry("keep.txt").Open()) WriteAll(keep, Encoding.UTF8.GetBytes("keep"));
            }
            package = AddEntryExtraFields(stream.ToArray(), "keep.txt", sourceExtraField, sourceExtraField);
        }

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip");
        int centralHeader = FindSignature(result.ToArray(), 0x02014B50u, "keep.txt");

        Assert.Equal(retainedExtraField, ReadLocalExtraField(result.ToArray(), centralHeader));
        Assert.Equal(retainedExtraField, ReadCentralExtraField(result.ToArray(), centralHeader));
    }

    [Fact]
    public void ZipRewriteDropsCommentDependentUnicodeCommentExtras() {
        byte[] retainedExtraField = { 0xFE, 0xCA, 0x01, 0x00, 0x42 };
        byte[] unicodeComment = Encoding.UTF8.GetBytes("legacy comment");
        byte[] unicodeCommentExtraField = new byte[9 + unicodeComment.Length];
        WriteLittleEndian16(unicodeCommentExtraField, 0, 0x6375);
        WriteLittleEndian16(unicodeCommentExtraField, 2, checked((ushort)(5 + unicodeComment.Length)));
        unicodeCommentExtraField[4] = 1;
        WriteLittleEndian(unicodeCommentExtraField, 5, 0xDEADBEEFu);
        Buffer.BlockCopy(unicodeComment, 0, unicodeCommentExtraField, 9, unicodeComment.Length);
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                using (Stream manifest = archive.CreateEntry("META-INF/content_credential.c2pa").Open()) WriteAll(manifest, CreateManifestStore());
                using (Stream keep = archive.CreateEntry("keep.txt").Open()) WriteAll(keep, Encoding.UTF8.GetBytes("keep"));
            }
            package = AddCentralDirectoryComment(stream.ToArray(), "keep.txt", new byte[] { 0x82 });
        }
        package = AddEntryExtraFields(package, "keep.txt", retainedExtraField, Join(retainedExtraField, unicodeCommentExtraField));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip");
        int centralHeader = FindSignature(result.ToArray(), 0x02014B50u, "keep.txt");

        Assert.Equal(retainedExtraField, ReadLocalExtraField(result.ToArray(), centralHeader));
        Assert.Equal(retainedExtraField, ReadCentralExtraField(result.ToArray(), centralHeader));
        Assert.Equal("é", Encoding.UTF8.GetString(ReadCentralDirectoryComment(result.ToArray(), centralHeader)));
    }

    [Fact]
    public void ZipRewriteBoundsExpandedLegacyCommentsAtTheFieldLimit() {
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                using (Stream manifest = archive.CreateEntry("META-INF/content_credential.c2pa").Open()) WriteAll(manifest, CreateManifestStore());
                using (Stream keep = archive.CreateEntry("keep.txt").Open()) WriteAll(keep, Encoding.UTF8.GetBytes("keep"));
            }
            package = AddCentralDirectoryComment(stream.ToArray(), "keep.txt", Enumerable.Repeat((byte)0x82, ushort.MaxValue).ToArray());
        }
        int sourceCentralHeader = FindSignature(package, 0x02014B50u, "keep.txt");
        WriteLittleEndian16(package, sourceCentralHeader + 8,
            (ushort)(BitConverter.ToUInt16(package, sourceCentralHeader + 8) & ~0x0800));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceRemover.Remove(package, "fixture.zip"));

        Assert.Contains("cannot be represented completely", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OpcSignatureDetectionIgnoresDirectoryAndUnrelatedXmlSignatureEntries() {
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteZipEntry(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
                WriteZipEntry(archive, "_xmlsignatures/", string.Empty);
                WriteZipEntry(archive, "_xmlsignatures/readme.txt", "not a signature");
            }
            package = stream.ToArray();
        }

        Assert.False(OfficeProvenanceZip.HasPackageSignature(package, new OfficeProvenanceRemovalOptions()));
    }

    [Fact]
    public void ZipRewriteResolvesAForcedZip64LocalHeaderOffset() {
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                ZipArchiveEntry manifest = archive.CreateEntry("META-INF/content_credential.c2pa");
                using (Stream target = manifest.Open()) WriteAll(target, CreateManifestStore());
                using Stream keep = archive.CreateEntry("keep.txt").Open();
                WriteAll(keep, Encoding.UTF8.GetBytes("keep"));
            }
            package = stream.ToArray();
        }
        int originalCentralHeader = FindSignature(package, 0x02014B50u, "keep.txt");
        uint localHeaderOffset = BitConverter.ToUInt32(package, originalCentralHeader + 42);
        byte[] zip64Extra = new byte[12];
        WriteLittleEndian16(zip64Extra, 0, 0x0001);
        WriteLittleEndian16(zip64Extra, 2, 8);
        WriteLittleEndian64(zip64Extra, 4, localHeaderOffset);
        package = AddEntryExtraFields(package, "keep.txt", Array.Empty<byte>(), zip64Extra);
        int centralHeader = FindSignature(package, 0x02014B50u, "keep.txt");
        WriteLittleEndian(package, centralHeader + 42, uint.MaxValue);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip");

        Assert.Equal("keep", Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "keep.txt")));
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void ZipRewriteResolvesAForcedZip64CentralDirectoryOffset() {
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                using (Stream manifest = archive.CreateEntry("META-INF/content_credential.c2pa").Open()) WriteAll(manifest, CreateManifestStore());
                using (Stream keep = archive.CreateEntry("keep.txt").Open()) WriteAll(keep, Encoding.UTF8.GetBytes("keep"));
            }
            package = PromoteToZip64CentralDirectoryOffset(stream.ToArray());
        }

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip");

        Assert.Equal("keep", Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "keep.txt")));
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void BigTiffHonorsTheConfiguredMaximumEntryBoundary() {
        byte[] tiff = CreateBigTiff(65535);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(
            tiff,
            "fixture.tiff",
            new OfficeProvenanceOptions { MaxContainerEntries = 65536 });

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void VerificationResultSnapshotsMutableFindings() {
        var findings = new List<string> { "initial" };
        var result = new OfficeProvenanceVerificationResult(
            OfficeProvenanceVerificationStatus.Valid, "test", findings);

        findings[0] = "changed";
        findings.Add("added");

        Assert.Equal(new[] { "initial" }, result.Findings);
    }

    [Fact]
    public void SignatureDiscoveryEnforcesAggregatePartBytesWhenDigestVerificationIsDisabled() {
        string contentTypes =
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Override PartName=\"/_xmlsignatures/sig1.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "<Override PartName=\"/_xmlsignatures/sig2.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "</Types>";
        string signature =
            "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><Object>" +
            new string('x', 2048) + "</Object></Signature>";
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteZipEntry(archive, "[Content_Types].xml", contentTypes);
                WriteZipEntry(archive, "_xmlsignatures/sig1.xml", signature);
                WriteZipEntry(archive, "_xmlsignatures/sig2.xml", signature);
            }
            package = output.ToArray();
        }
        int signatureBytes = Encoding.UTF8.GetByteCount(signature);
        var bounded = new OfficePackageSignatureInspectionOptions {
            VerifyDigests = false,
            MaxSignatureBytes = signatureBytes + 1L,
            MaxTotalDigestBytes = signatureBytes + 1L
        };

        OfficePackageSignatureInfo rejected = OfficePackageSignatureService.Inspect(package, bounded);
        OfficePackageSignatureInfo accepted = OfficePackageSignatureService.Inspect(package,
            new OfficePackageSignatureInspectionOptions {
                VerifyDigests = false,
                MaxSignatureBytes = signatureBytes + 1L,
                MaxTotalDigestBytes = signatureBytes * 2L
            });

        Assert.Contains(rejected.SignatureParts, part =>
            part.ParseError?.Contains("aggregate limit", StringComparison.OrdinalIgnoreCase) == true);
        Assert.DoesNotContain(accepted.SignatureParts, part =>
            part.ParseError?.Contains("aggregate limit", StringComparison.OrdinalIgnoreCase) == true);
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void IncompleteExtendedXmpDoesNotAllocateTheDeclaredPacketLength() {
        const string guid = "0123456789ABCDEF0123456789ABCDEF";
        byte[] standardPacket = Encoding.UTF8.GetBytes(
            $"<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:xmpNote=\"http://ns.adobe.com/xmp/note/\" xmpNote:HasExtendedXMP=\"{guid}\"/>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), standardPacket)),
            CreateExtendedXmpSegment(guid, new byte[] { 1 }, 128 * 1024 * 1024),
            new byte[] { 0xFF, 0xD9 });

        long before = GC.GetAllocatedBytesForCurrentThread();
        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.Empty(report.Evidence);
        Assert.True(allocated < 8L * 1024L * 1024L, $"Inspection allocated {allocated} bytes.");
    }


    [Fact]
    public void IncompleteApp11SequenceDoesNotAllocateTheDeclaredManifestLength() {
        byte[] fragment = CreateManifestStore();
        WriteBigEndian(fragment, 0, 64 * 1024 * 1024);
        byte[] app11Payload = Join(
            Encoding.ASCII.GetBytes("JP"),
            new byte[] { 0x12, 0x34 },
            BigEndian(1),
            fragment);
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xEB, app11Payload),
            new byte[] { 0xFF, 0xD9 });

        long before = GC.GetAllocatedBytesForCurrentThread();
        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.True(allocated < 8L * 1024L * 1024L, $"Inspection allocated {allocated} bytes.");
    }
#endif

    [Fact]
    public void ZipEmbeddedAssetsShareTheTopLevelCarrierLimit() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                using (Stream first = archive.CreateEntry("media/first.png").Open()) WriteAll(first, image);
                using (Stream second = archive.CreateEntry("media/second.png").Open()) WriteAll(second, image);
            }
            package = stream.ToArray();
        }
        var inspectionOptions = new OfficeProvenanceOptions { MaxCarriers = 1 };
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxCarriers = 1;

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(package, "fixture.zip", inspectionOptions));
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceRemover.Remove(package, "fixture.zip", removalOptions));
    }

    [Fact]
    public void TiffIfdEntriesShareTheConfiguredContainerEntryLimit() {
        byte[] tiff = CreateTiffWithTwoIfds(entriesPerIfd: 2);
        var options = new OfficeProvenanceOptions { MaxContainerEntries = 3 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceInspector.Inspect(tiff, "fixture.tiff", options));

        Assert.Contains("container-entry limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TiffMainIfdCountUsesTheConfiguredContainerEntryLimit() {
        byte[] tiff = CreateTiffWithEmptyIfdChain(1025);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(
            tiff,
            "fixture.tiff",
            new OfficeProvenanceOptions { MaxContainerEntries = 1025 });

        Assert.Empty(report.Evidence);
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            tiff,
            "fixture.tiff",
            new OfficeProvenanceOptions { MaxContainerEntries = 1024 }));
    }

    [Fact]
    public void TiffRejectsOutOfLinePayloadCountsLargerThanTheAsset() {
        byte[] tiff = CreateTiffWithDeclaredXmpPayloadLength(1024 * 1024);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(tiff, "fixture.tiff");

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void TiffPreservesRepeatedTagsThatShareTheSameXmpPayloadRange() {
        byte[] xmp = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
            "<rdf:Description xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\" " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
            "</rdf:RDF></x:xmpmeta>");
        byte[] tiff = CreateTiffWithRepeatedXmpRangeAcrossIfds(xmp, 3);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(
            tiff,
            "fixture.tiff",
            new OfficeProvenanceOptions { MaxExpandedContainerBytes = xmp.Length });

        Assert.Single(report.Evidence);

        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxExpandedContainerBytes = xmp.Length;
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tiff", removalOptions);

        Assert.False(result.WasChanged);
        Assert.True(result.After.HasGenerativeAiDeclaration);
        Assert.Equal(tiff, result.ToArray());
    }

    [Fact]
    public void SvgPreservesMixedWrappedAndDirectIptcScopesAsAmbiguous() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:x=\"adobe:ns:meta/\" " +
            "xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" +
            "<metadata><x:xmpmeta><rdf:RDF><rdf:Description iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></rdf:RDF></x:xmpmeta></metadata>" +
            "<metadata><direct iptc:DigitalSourceType=\"preserve-non-xmp\"/></metadata>" +
            "<rect width=\"1\" height=\"1\"/></svg>");

        OfficeProvenanceReport before = OfficeProvenanceInspector.Inspect(svg, "fixture.svg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Single(before.Evidence);
        Assert.False(Assert.Single(before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Single(result.After.Evidence);
        Assert.Contains("preserve-non-xmp", Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void SvgProcessesOnlyTheOutermostNestedXmpRoot() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:x=\"adobe:ns:meta/\" " +
            "xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><metadata><x:xmpmeta>" +
            "<x:xmpmeta><rdf:RDF><rdf:Description iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
            "</rdf:RDF></x:xmpmeta></x:xmpmeta></metadata></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Single(result.Before.Evidence);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void GifProcessesStandardXmpApplicationExtensions() {
        byte[] xmp = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
            "<rdf:Description xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\" " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
            "</rdf:RDF></x:xmpmeta>");
        byte[] gif = CreateGifWithXmp(xmp);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(gif, "fixture.gif");

        Assert.Single(result.Before.Evidence);
        Assert.True(result.WasChanged);
        Assert.True(result.WasReserialized);
        Assert.Empty(result.After.Evidence);
        Assert.Equal((byte)0x3B, result.ToArray()[result.ToArray().Length - 1]);
    }

    [Fact]
    public void SvgRejectsAValidDocumentBeforeMaterializingTooManyNodes() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\"><g/><g/><g/><g/><g/></svg>");
        var bounded = new OfficeProvenanceOptions { MaxContainerEntries = 5 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceInspector.Inspect(svg, "fixture.svg", bounded));

        Assert.Contains("XML node limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(OfficeProvenanceInspector.Inspect(
            svg,
            "fixture.svg",
            new OfficeProvenanceOptions { MaxContainerEntries = 16 }).Evidence);
    }

    [Fact]
    public void ExtensionlessSvgSniffingUsesTheConfiguredXmlNodeLimit() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<!--first--><!--second--><svg xmlns=\"http://www.w3.org/2000/svg\"/>");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceInspector.Inspect(
                svg,
                "fixture.bin",
                new OfficeProvenanceOptions { MaxContainerEntries = 3 }));

        Assert.Contains("XML node limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(
            OfficeProvenanceAssetFormat.Svg,
            OfficeProvenanceInspector.Inspect(
                svg,
                "fixture.bin",
                new OfficeProvenanceOptions { MaxContainerEntries = 4 }).Format);
    }

    [Fact]
    public void RemovalResultSnapshotsMutableConstructorInputs() {
        byte[] data = { 1, 2, 3 };
        var changes = new List<OfficeProvenanceChange> {
            new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, "test", 1)
        };
        var report = new OfficeProvenanceReport(OfficeProvenanceAssetFormat.Unknown, Array.Empty<OfficeProvenanceEvidence>());
        var result = new OfficeProvenanceRemovalResult(data, report, report, changes, wasReserialized: false);

        data[0] = 9;
        changes.Clear();

        Assert.Equal(new byte[] { 1, 2, 3 }, result.ToArray());
        Assert.True(result.WasChanged);
        Assert.Single(result.Changes);
    }

    private static void WriteAll(Stream stream, byte[] data) => stream.Write(data, 0, data.Length);

    private static void WriteZipEntry(ZipArchive archive, string name, string content) {
        using Stream stream = archive.CreateEntry(name, CompressionLevel.Optimal).Open();
        WriteAll(stream, Encoding.UTF8.GetBytes(content));
    }

    private static byte[] CreateManifestStore(bool includeSignature = true) {
        byte[] uuid = { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] descriptionPayload = Join(uuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa\0"));
        byte[] description = CreateBox("jumd", descriptionPayload);
        byte[] manifestUuid = { 0x63, 0x32, 0x6D, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] manifestDescription = CreateBox("jumd", Join(manifestUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("m\0")));
        byte[] assertionStoreUuid = { 0x63, 0x32, 0x61, 0x73, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] assertionStoreDescription = CreateBox("jumd", Join(assertionStoreUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.assertions\0")));
        byte[] assertionUuid = { 0x63, 0x32, 0x61, 0x63, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] assertionDescription = CreateBox("jumd", Join(assertionUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.test\0")));
        byte[] assertionStore = CreateBox("jumb", Join(assertionStoreDescription,
            CreateBox("jumb", Join(assertionDescription, CreateBox("cbor", new byte[] { 0xA0 })))));
        byte[] claimUuid = { 0x63, 0x32, 0x63, 0x6C, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] claimDescription = CreateBox("jumd", Join(claimUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.claim\0")));
        byte[] claim = CreateBox("jumb", Join(claimDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] signatureUuid = { 0x63, 0x32, 0x63, 0x73, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] signatureDescription = CreateBox("jumd", Join(signatureUuid, new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.signature\0")));
        byte[] signature = CreateBox("jumb", Join(signatureDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] manifest = includeSignature
            ? CreateBox("jumb", Join(manifestDescription, assertionStore, claim, signature))
            : CreateBox("jumb", Join(manifestDescription, assertionStore, claim));
        return CreateBox("jumb", Join(description, manifest));
    }

    private static byte[] CreateBox(string type, byte[] payload) {
        byte[] box = new byte[8 + payload.Length];
        WriteBigEndian(box, 0, box.Length);
        Encoding.ASCII.GetBytes(type, 0, 4, box, 4);
        Buffer.BlockCopy(payload, 0, box, 8, payload.Length);
        return box;
    }

    private static byte[] ReadZipEntry(byte[] package, string name) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        ZipArchiveEntry entry = archive.GetEntry(name) ?? throw new InvalidDataException("ZIP fixture entry was not found.");
        using Stream source = entry.Open();
        using var output = new MemoryStream();
        source.CopyTo(output);
        return output.ToArray();
    }

    private static byte[] CreateExtendedBox(string type, byte[] payload) {
        byte[] box = new byte[16 + payload.Length];
        WriteBigEndian(box, 0, 1);
        Encoding.ASCII.GetBytes(type, 0, 4, box, 4);
        WriteBigEndian64(box, 8, (ulong)box.Length);
        Buffer.BlockCopy(payload, 0, box, 16, payload.Length);
        return box;
    }

    private static byte[] CreatePngWithManifest(byte[] manifest) => Join(
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
        CreatePngChunk("IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 }),
        CreatePngChunk("caBX", manifest),
        CreatePngChunk("IDAT", new byte[] { 0x78, 0x9C, 0x63, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x04, 0x00, 0x01 }),
        CreatePngChunk("IEND", Array.Empty<byte>()));

    private static byte[] CreateWebp(params byte[][] chunks) {
        byte[] result = Join(new byte[12], Join(chunks));
        Encoding.ASCII.GetBytes("RIFF").CopyTo(result, 0);
        WriteLittleEndian(result, 4, (uint)(result.Length - 8));
        Encoding.ASCII.GetBytes("WEBP").CopyTo(result, 8);
        return result;
    }

    private static byte[] CreateGifWithXmp(byte[] packet) {
        byte[] trailer = new byte[258];
        trailer[0] = 0x01;
        for (int index = 1; index <= 255; index++) trailer[index] = checked((byte)(256 - index));
        return Join(
            Encoding.ASCII.GetBytes("GIF89a"),
            new byte[] { 1, 0, 1, 0, 0, 0, 0 },
            new byte[] { 0x21, 0xFF, 0x0B },
            Encoding.ASCII.GetBytes("XMP DataXMP"),
            packet,
            trailer,
            CreateMinimalGifImage(),
            new byte[] { 0x3B });
    }

    private static byte[] CreateRiffChunk(string type, byte[] payload) {
        byte[] chunk = new byte[8 + payload.Length + (payload.Length & 1)];
        Encoding.ASCII.GetBytes(type).CopyTo(chunk, 0);
        WriteLittleEndian(chunk, 4, (uint)payload.Length);
        Buffer.BlockCopy(payload, 0, chunk, 8, payload.Length);
        return chunk;
    }

    private static byte[] CreatePngChunk(string type, byte[] payload) {
        byte[] typeBytes = Encoding.ASCII.GetBytes(type);
        byte[] chunk = new byte[12 + payload.Length];
        WriteBigEndian(chunk, 0, payload.Length);
        Buffer.BlockCopy(typeBytes, 0, chunk, 4, 4);
        Buffer.BlockCopy(payload, 0, chunk, 8, payload.Length);
        WriteBigEndian(chunk, 8 + payload.Length, unchecked((int)Crc32(Join(typeBytes, payload))));
        return chunk;
    }

    private static uint Crc32(byte[] data) {
        uint crc = 0xFFFFFFFF;
        foreach (byte value in data) {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++) crc = (crc >> 1) ^ (0xEDB88320U & (uint)-(int)(crc & 1));
        }
        return ~crc;
    }

    private static byte[] CreateTextWrapper(byte[] manifest) {
        byte[] header = Join(Encoding.ASCII.GetBytes("C2PATXT\0"), new byte[] { 1 }, BigEndian(manifest.Length), manifest);
        var builder = new StringBuilder("\uFEFF");
        foreach (byte value in header) builder.Append(char.ConvertFromUtf32(value < 16 ? 0xFE00 + value : 0xE0100 + value - 16));
        return Encoding.UTF8.GetBytes(builder.ToString());
    }

    private static byte[] CreateExtendedXmpSegment(string guid, byte[] packet) {
        return CreateExtendedXmpSegment(guid, packet, packet.Length);
    }

    private static byte[] CreateExtendedXmpSegment(string guid, byte[] packet, int declaredLength) {
        byte[] header = Encoding.ASCII.GetBytes("http://ns.adobe.com/xmp/extension/\0");
        byte[] payload = new byte[header.Length + 40 + packet.Length];
        Buffer.BlockCopy(header, 0, payload, 0, header.Length);
        Encoding.ASCII.GetBytes(guid, 0, guid.Length, payload, header.Length);
        WriteBigEndian(payload, header.Length + 32, declaredLength);
        WriteBigEndian(payload, header.Length + 36, 0);
        Buffer.BlockCopy(packet, 0, payload, header.Length + 40, packet.Length);
        return CreateJpegSegment(0xE1, payload);
    }

    private static int FindSignature(byte[] data, uint signature, string entryName) {
        byte[] name = Encoding.UTF8.GetBytes(entryName);
        for (int index = 0; index + 46 + name.Length <= data.Length; index++) {
            if (BitConverter.ToUInt32(data, index) != signature) continue;
            int nameLength = BitConverter.ToUInt16(data, index + 28);
            if (nameLength == name.Length && data.AsSpan(index + 46, nameLength).SequenceEqual(name)) return index;
        }
        throw new InvalidDataException("ZIP central-directory entry was not found.");
    }

    private static byte[] AddCentralDirectoryComment(byte[] package, string entryName, byte[] comment) {
        int centralHeader = FindSignature(package, 0x02014B50u, entryName);
        int nameLength = BitConverter.ToUInt16(package, centralHeader + 28);
        int extraLength = BitConverter.ToUInt16(package, centralHeader + 30);
        Assert.Equal(0, BitConverter.ToUInt16(package, centralHeader + 32));
        int insertOffset = centralHeader + 46 + nameLength + extraLength;
        int endOffset = -1;
        for (int index = package.Length - 22; index >= 0; index--) {
            if (BitConverter.ToUInt32(package, index) == 0x06054B50u) { endOffset = index; break; }
        }
        if (endOffset < 0) throw new InvalidDataException("ZIP end record was not found.");
        byte[] updated = new byte[package.Length + comment.Length];
        Buffer.BlockCopy(package, 0, updated, 0, insertOffset);
        Buffer.BlockCopy(comment, 0, updated, insertOffset, comment.Length);
        Buffer.BlockCopy(package, insertOffset, updated, insertOffset + comment.Length, package.Length - insertOffset);
        WriteLittleEndian16(updated, centralHeader + 32, checked((ushort)comment.Length));
        int updatedEndOffset = endOffset + comment.Length;
        uint centralSize = BitConverter.ToUInt32(updated, updatedEndOffset + 12);
        WriteLittleEndian(updated, updatedEndOffset + 12, checked(centralSize + (uint)comment.Length));
        return updated;
    }

    private static byte[] AddEntryExtraFields(
        byte[] package,
        string entryName,
        byte[] localExtraField,
        byte[] centralExtraField) {
        int centralHeader = FindSignature(package, 0x02014B50u, entryName);
        int localHeader = checked((int)BitConverter.ToUInt32(package, centralHeader + 42));
        Assert.Equal(0x04034B50u, BitConverter.ToUInt32(package, localHeader));
        int localNameLength = BitConverter.ToUInt16(package, localHeader + 26);
        Assert.Equal(0, BitConverter.ToUInt16(package, localHeader + 28));
        int centralNameLength = BitConverter.ToUInt16(package, centralHeader + 28);
        Assert.Equal(0, BitConverter.ToUInt16(package, centralHeader + 30));
        int localInsertOffset = localHeader + 30 + localNameLength;
        int centralInsertOffset = centralHeader + 46 + centralNameLength;
        int endOffset = FindEndOfCentralDirectory(package);
        byte[] updated = new byte[package.Length + localExtraField.Length + centralExtraField.Length];
        Buffer.BlockCopy(package, 0, updated, 0, localInsertOffset);
        Buffer.BlockCopy(localExtraField, 0, updated, localInsertOffset, localExtraField.Length);
        Buffer.BlockCopy(package, localInsertOffset, updated, localInsertOffset + localExtraField.Length, centralInsertOffset - localInsertOffset);
        Buffer.BlockCopy(centralExtraField, 0, updated, centralInsertOffset + localExtraField.Length, centralExtraField.Length);
        Buffer.BlockCopy(package, centralInsertOffset, updated, centralInsertOffset + localExtraField.Length + centralExtraField.Length, package.Length - centralInsertOffset);
        WriteLittleEndian16(updated, localHeader + 28, checked((ushort)localExtraField.Length));
        int updatedCentralHeader = centralHeader + localExtraField.Length;
        WriteLittleEndian16(updated, updatedCentralHeader + 30, checked((ushort)centralExtraField.Length));
        int updatedEndOffset = endOffset + localExtraField.Length + centralExtraField.Length;
        uint centralSize = BitConverter.ToUInt32(updated, updatedEndOffset + 12);
        uint centralOffset = BitConverter.ToUInt32(updated, updatedEndOffset + 16);
        WriteLittleEndian(updated, updatedEndOffset + 12, checked(centralSize + (uint)centralExtraField.Length));
        WriteLittleEndian(updated, updatedEndOffset + 16, checked(centralOffset + (uint)localExtraField.Length));
        return updated;
    }

    private static byte[] ReadLocalExtraField(byte[] package, int centralHeader) {
        int localHeader = checked((int)BitConverter.ToUInt32(package, centralHeader + 42));
        int nameLength = BitConverter.ToUInt16(package, localHeader + 26);
        int extraLength = BitConverter.ToUInt16(package, localHeader + 28);
        return package.AsSpan(localHeader + 30 + nameLength, extraLength).ToArray();
    }

    private static byte[] ReadCentralExtraField(byte[] package, int centralHeader) {
        int nameLength = BitConverter.ToUInt16(package, centralHeader + 28);
        int extraLength = BitConverter.ToUInt16(package, centralHeader + 30);
        return package.AsSpan(centralHeader + 46 + nameLength, extraLength).ToArray();
    }

    private static byte[] ReadCentralDirectoryComment(byte[] package, int centralHeader) {
        int nameLength = BitConverter.ToUInt16(package, centralHeader + 28);
        int extraLength = BitConverter.ToUInt16(package, centralHeader + 30);
        int commentLength = BitConverter.ToUInt16(package, centralHeader + 32);
        byte[] comment = new byte[commentLength];
        Buffer.BlockCopy(package, centralHeader + 46 + nameLength + extraLength, comment, 0, commentLength);
        return comment;
    }

    private static byte[] AddArchiveComment(byte[] package, byte[] comment) {
        int endOffset = FindEndOfCentralDirectory(package);
        Assert.Equal(0, BitConverter.ToUInt16(package, endOffset + 20));
        byte[] updated = new byte[package.Length + comment.Length];
        Buffer.BlockCopy(package, 0, updated, 0, package.Length);
        WriteLittleEndian16(updated, endOffset + 20, checked((ushort)comment.Length));
        Buffer.BlockCopy(comment, 0, updated, package.Length, comment.Length);
        return updated;
    }

    private static byte[] ReadArchiveComment(byte[] package) {
        int endOffset = FindEndOfCentralDirectory(package);
        int length = BitConverter.ToUInt16(package, endOffset + 20);
        return package.AsSpan(endOffset + 22, length).ToArray();
    }

    private static int FindEndOfCentralDirectory(byte[] package) {
        for (int index = package.Length - 22; index >= Math.Max(0, package.Length - 22 - ushort.MaxValue); index--) {
            if (BitConverter.ToUInt32(package, index) == 0x06054B50u &&
                index + 22 + BitConverter.ToUInt16(package, index + 20) == package.Length) return index;
        }
        throw new InvalidDataException("ZIP end record was not found.");
    }

    private static byte[] CreateJpegSegment(byte marker, byte[] payload) {
        byte[] segment = new byte[payload.Length + 4];
        segment[0] = 0xFF;
        segment[1] = marker;
        int length = payload.Length + 2;
        segment[2] = (byte)(length >> 8);
        segment[3] = (byte)length;
        Buffer.BlockCopy(payload, 0, segment, 4, payload.Length);
        return segment;
    }

    private static byte[] CreateMinimalJpegFrame() => CreateJpegSegment(
        0xC0,
        new byte[] { 8, 0, 1, 0, 1, 1, 1, 0x11, 0 });

    private static byte[] CreateMinimalJpegScan() => CreateJpegSegment(
        0xDA,
        new byte[] { 1, 1, 0, 0, 63, 0 });

    private static byte[] CreateMinimalGifImage() =>
        new byte[] {
            0x2C, 0, 0, 0, 0, 1, 0, 1, 0, 0x80,
            0, 0, 0, 255, 255, 255,
            2, 2, 0x44, 0x01, 0
        };

    private static byte[] CreateJpegApp11(byte[] manifest, int offset, int count, ushort instance, uint sequence) {
        byte[] payload = new byte[count + 8];
        payload[0] = 0x4A;
        payload[1] = 0x50;
        payload[2] = (byte)(instance >> 8);
        payload[3] = (byte)instance;
        WriteBigEndian(payload, 4, checked((int)sequence));
        Buffer.BlockCopy(manifest, offset, payload, 8, count);
        return CreateJpegSegment(0xEB, payload);
    }

    private static byte[] CreateValidJpeg(params byte[][] segments) {
        byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        return Join(jpeg.Take(2).ToArray(), Join(segments), jpeg.Skip(2).ToArray());
    }

    private static string ComputeMd5(byte[] data) {
        using MD5 md5 = MD5.Create();
        return string.Concat(md5.ComputeHash(data).Select(value => value.ToString("X2")));
    }

    private static byte[] CreateZip64CountOnlyPackage(ulong count) {
        byte[] package = new byte[102];
        package[0] = 0x50; package[1] = 0x4B; package[2] = 0x03; package[3] = 0x04;
        WriteLittleEndian(package, 4, 0x06064B50U);
        WriteLittleEndian64(package, 8, 44);
        WriteLittleEndian64(package, 28, count);
        WriteLittleEndian64(package, 36, count);
        WriteLittleEndian(package, 60, 0x07064B50U);
        WriteLittleEndian64(package, 68, 4);
        WriteLittleEndian(package, 76, 1U);
        WriteLittleEndian(package, 80, 0x06054B50U);
        package[88] = 0xFF; package[89] = 0xFF;
        package[90] = 0xFF; package[91] = 0xFF;
        return package;
    }

    private static byte[] PromoteToZip64CentralDirectoryOffset(byte[] package) {
        const int zip64RecordLength = 56;
        const int zip64LocatorLength = 20;
        int endOffset = FindEndOfCentralDirectory(package);
        ushort entries = BitConverter.ToUInt16(package, endOffset + 10);
        uint centralSize = BitConverter.ToUInt32(package, endOffset + 12);
        uint centralOffset = BitConverter.ToUInt32(package, endOffset + 16);
        byte[] result = new byte[package.Length + zip64RecordLength + zip64LocatorLength];
        Buffer.BlockCopy(package, 0, result, 0, endOffset);
        int zip64Offset = endOffset;
        WriteLittleEndian(result, zip64Offset, 0x06064B50U);
        WriteLittleEndian64(result, zip64Offset + 4, 44);
        WriteLittleEndian16(result, zip64Offset + 12, 45);
        WriteLittleEndian16(result, zip64Offset + 14, 45);
        WriteLittleEndian64(result, zip64Offset + 24, entries);
        WriteLittleEndian64(result, zip64Offset + 32, entries);
        WriteLittleEndian64(result, zip64Offset + 40, centralSize);
        WriteLittleEndian64(result, zip64Offset + 48, centralOffset);
        int locatorOffset = zip64Offset + zip64RecordLength;
        WriteLittleEndian(result, locatorOffset, 0x07064B50U);
        WriteLittleEndian64(result, locatorOffset + 8, (ulong)zip64Offset);
        WriteLittleEndian(result, locatorOffset + 16, 1);
        int updatedEndOffset = locatorOffset + zip64LocatorLength;
        Buffer.BlockCopy(package, endOffset, result, updatedEndOffset, package.Length - endOffset);
        WriteLittleEndian(result, updatedEndOffset + 16, uint.MaxValue);
        return result;
    }

    private static byte[] CreateTiffWithTwoIfds(int entriesPerIfd) {
        int ifdSize = 2 + entriesPerIfd * 12 + 4;
        byte[] data = new byte[8 + ifdSize * 2];
        data[0] = (byte)'I';
        data[1] = (byte)'I';
        data[2] = 42;
        WriteLittleEndian(data, 4, 8U);
        WriteLittleEndian16(data, 8, (ushort)entriesPerIfd);
        WriteLittleEndian(data, 8 + 2 + entriesPerIfd * 12, (uint)(8 + ifdSize));
        WriteLittleEndian16(data, 8 + ifdSize, (ushort)entriesPerIfd);
        return data;
    }

    private static byte[] CreateTiffWithEmptyIfdChain(int ifdCount) {
        byte[] data = new byte[8 + ifdCount * 6];
        data[0] = data[1] = (byte)'I';
        data[2] = 42;
        WriteLittleEndian(data, 4, 8U);
        for (int index = 0; index < ifdCount; index++) {
            int ifdOffset = 8 + index * 6;
            uint nextOffset = index + 1 < ifdCount ? (uint)(ifdOffset + 6) : 0U;
            WriteLittleEndian(data, ifdOffset + 2, nextOffset);
        }
        return data;
    }

    private static byte[] CreateTiffWithDeclaredXmpPayloadLength(int payloadLength) {
        byte[] data = new byte[26];
        data[0] = data[1] = (byte)'I';
        data[2] = 42;
        WriteLittleEndian(data, 4, 8U);
        WriteLittleEndian16(data, 8, 1);
        WriteLittleEndian16(data, 10, 700);
        WriteLittleEndian16(data, 12, 1);
        WriteLittleEndian(data, 14, checked((uint)payloadLength));
        WriteLittleEndian(data, 18, 26U);
        return data;
    }

    private static byte[] CreateTiffWithRepeatedXmpRangeAcrossIfds(byte[] xmp, int ifdCount) {
        const int firstIfdOffset = 8;
        const int ifdSize = 18;
        int payloadOffset = firstIfdOffset + ifdCount * ifdSize;
        byte[] result = new byte[payloadOffset + xmp.Length];
        result[0] = result[1] = (byte)'I';
        result[2] = 42;
        BitConverter.GetBytes(firstIfdOffset).CopyTo(result, 4);
        for (int index = 0; index < ifdCount; index++) {
            int ifdOffset = firstIfdOffset + index * ifdSize;
            BitConverter.GetBytes((ushort)1).CopyTo(result, ifdOffset);
            int entryOffset = ifdOffset + 2;
            BitConverter.GetBytes((ushort)700).CopyTo(result, entryOffset);
            BitConverter.GetBytes((ushort)1).CopyTo(result, entryOffset + 2);
            BitConverter.GetBytes(xmp.Length).CopyTo(result, entryOffset + 4);
            BitConverter.GetBytes(payloadOffset).CopyTo(result, entryOffset + 8);
            int nextFieldOffset = entryOffset + 12;
            int nextIfdOffset = index + 1 < ifdCount ? ifdOffset + ifdSize : 0;
            BitConverter.GetBytes(nextIfdOffset).CopyTo(result, nextFieldOffset);
        }
        Buffer.BlockCopy(xmp, 0, result, payloadOffset, xmp.Length);
        return result;
    }

    private static byte[] CreateBigTiff(int entryCount) {
        byte[] data = new byte[16 + 8 + checked(entryCount * 20) + 8];
        data[0] = (byte)'I';
        data[1] = (byte)'I';
        WriteLittleEndian16(data, 2, 43);
        WriteLittleEndian16(data, 4, 8);
        WriteLittleEndian64(data, 8, 16);
        WriteLittleEndian64(data, 16, (ulong)entryCount);
        return data;
    }

    private static byte[] BigEndian(int value) {
        byte[] bytes = new byte[4];
        WriteBigEndian(bytes, 0, value);
        return bytes;
    }

    private static void WriteBigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static void WriteLittleEndian(byte[] data, int offset, uint value) {
        data[offset] = (byte)value;
        data[offset + 1] = (byte)(value >> 8);
        data[offset + 2] = (byte)(value >> 16);
        data[offset + 3] = (byte)(value >> 24);
    }

    private static void WriteLittleEndian16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)value;
        data[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteLittleEndian64(byte[] data, int offset, ulong value) {
        for (int index = 0; index < 8; index++) data[offset + index] = (byte)(value >> (index * 8));
    }

    private static void WriteBigEndian64(byte[] data, int offset, ulong value) {
        for (int index = 0; index < 8; index++) data[offset + index] = (byte)(value >> ((7 - index) * 8));
    }

    private static byte[] Join(params byte[][] values) {
        byte[] result = new byte[values.Sum(value => value.Length)];
        int offset = 0;
        foreach (byte[] value in values) {
            Buffer.BlockCopy(value, 0, result, offset, value.Length);
            offset += value.Length;
        }
        return result;
    }
}

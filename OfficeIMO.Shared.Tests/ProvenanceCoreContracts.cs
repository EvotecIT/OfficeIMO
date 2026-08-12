using System.IO.Compression;
using System.Text;
using OfficeIMO;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ProvenanceCoreContracts {
    [Fact]
    public void JpegRemovesExactApp11SequenceAndPreservesOtherSegments() {
        byte[] manifest = CreateManifestStore(64);
        byte[] unrelated = CreateJpegSegment(0xEB, Encoding.ASCII.GetBytes("not-c2pa"));
        byte[] first = CreateJpegApp11(manifest, 0, 40, instance: 7, sequence: 1);
        byte[] second = CreateJpegApp11(manifest, 40, manifest.Length - 40, instance: 7, sequence: 2);
        byte[] jpeg = Join(new byte[] { 0xFF, 0xD8 }, unrelated, first, second, new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.Single(report.Evidence);
        Assert.True(report.Evidence[0].IsStructurallyValid);
        Assert.Equal(manifest.Length, report.Evidence[0].PayloadLength);
        Assert.Equal(Join(new byte[] { 0xFF, 0xD8 }, unrelated, new byte[] { 0xFF, 0xD9 }), result.ToArray());
        Assert.Empty(result.After.Evidence);
        Assert.False(result.WasReserialized);
    }

    [Fact]
    public void JpegPreservesMalformedOrNonContiguousApp11SequenceByDefault() {
        byte[] manifest = CreateManifestStore(64);
        byte[] first = CreateJpegApp11(manifest, 0, 40, instance: 7, sequence: 1);
        byte[] intervening = CreateJpegSegment(0xE1, Encoding.ASCII.GetBytes("unrelated"));
        byte[] second = CreateJpegApp11(manifest, 40, manifest.Length - 40, instance: 7, sequence: 2);
        byte[] jpeg = Join(new byte[] { 0xFF, 0xD8 }, first, intervening, second, new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.False(result.WasChanged);
        Assert.Equal(jpeg, result.ToArray());
    }

    [Fact]
    public void JpegCanRemoveRecognizableMalformedApp11SequenceWhenExplicitlyRequested() {
        byte[] malformedManifest = CreateManifestStore();
        WriteBigEndian(malformedManifest, 0, malformedManifest.Length + 10);
        byte[] app11 = CreateJpegApp11(malformedManifest, 0, malformedManifest.Length, instance: 7, sequence: 1);
        byte[] jpeg = Join(new byte[] { 0xFF, 0xD8 }, app11, new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult preserved = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult removed = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg", new OfficeProvenanceRemovalOptions {
            RequireStructurallyValidCarrier = false
        });

        Assert.Single(report.Evidence);
        Assert.False(report.Evidence[0].IsStructurallyValid);
        Assert.Equal(jpeg, preserved.ToArray());
        Assert.Equal(new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 }, removed.ToArray());
    }

    [Fact]
    public void JpegIgnoresGenericApp11JumbfWithoutTheC2paIdentity() {
        byte[] generic = CreateManifestStore();
        generic[16] ^= 0x01;
        byte[] app11 = CreateJpegApp11(generic, 0, generic.Length, instance: 7, sequence: 1);
        byte[] jpeg = Join(new byte[] { 0xFF, 0xD8 }, app11, new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg", new OfficeProvenanceRemovalOptions {
            RequireStructurallyValidCarrier = false
        });

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
        Assert.Equal(jpeg, result.ToArray());
    }

    [Fact]
    public void JpegAcceptsFragmentedExtendedLengthJumbf() {
        byte[] manifest = CreateExtendedManifestStore(80);
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegApp11(manifest, 0, 50, instance: 9, sequence: 1),
            CreateJpegApp11(manifest, 50, manifest.Length - 50, instance: 9, sequence: 2),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.Single(result.Before.Evidence);
        Assert.True(result.Before.Evidence[0].IsStructurallyValid);
        Assert.Equal(manifest.Length, result.Before.Evidence[0].PayloadLength);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void PngRemovesCabxAndPreservesEveryOtherChunkByteForByte() {
        byte[] manifest = CreateManifestStore();
        byte[] header = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] ihdr = CreatePngChunk("IHDR", new byte[13]);
        byte[] cabx = CreatePngChunk("caBX", manifest);
        byte[] text = CreatePngChunk("tEXt", Encoding.ASCII.GetBytes("keep-this"));
        byte[] iend = CreatePngChunk("IEND", Array.Empty<byte>());
        byte[] png = Join(header, ihdr, cabx, text, iend);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.Equal(Join(header, ihdr, text, iend), result.ToArray());
        Assert.Single(result.Changes);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void PngPreservesCabxWithAnInvalidCrcByDefault() {
        byte[] header = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] cabx = CreatePngChunk("caBX", CreateManifestStore());
        cabx[cabx.Length - 1] ^= 0x01;
        byte[] png = Join(header, cabx, CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.Single(result.Before.Evidence);
        Assert.False(result.Before.Evidence[0].IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void PngPreservesManifestStoreWithTrailingCarrierBytesByDefault() {
        byte[] header = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] cabx = CreatePngChunk("caBX", Join(CreateManifestStore(), new byte[] { 1, 2, 3 }));
        byte[] png = Join(header, cabx, CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.Single(result.Before.Evidence);
        Assert.False(result.Before.Evidence[0].IsStructurallyValid);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void PngPreservesCabxPlacedAfterImageDataByDefault() {
        byte[] header = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] png = Join(
            header,
            CreatePngChunk("IHDR", new byte[13]),
            CreatePngChunk("IDAT", Array.Empty<byte>()),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void WebpRemovesC2paChunkAndRecomputesRiffLength() {
        byte[] keep = CreateRiffChunk("VP8 ", new byte[] { 1, 2, 3 });
        byte[] c2pa = CreateRiffChunk("C2PA", CreateManifestStore());
        byte[] webp = CreateWebp(keep, c2pa);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");
        byte[] expected = CreateWebp(keep);

        Assert.Equal(expected, result.ToArray());
        Assert.Equal(expected.Length - 8, BitConverter.ToInt32(expected, 4));
    }

    [Fact]
    public void WebpRiffLengthExcludesPreservedSuffixAfterRemoval() {
        byte[] keep = CreateRiffChunk("VP8 ", new byte[] { 1, 2, 3 });
        byte[] suffix = Encoding.ASCII.GetBytes("suffix");
        byte[] webp = Join(CreateWebp(keep, CreateRiffChunk("C2PA", CreateManifestStore())), suffix);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");
        byte[] expectedContainer = CreateWebp(keep);
        byte[] output = result.ToArray();

        Assert.Equal(Join(expectedContainer, suffix), output);
        Assert.Equal(expectedContainer.Length - 8, BitConverter.ToInt32(output, 4));
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void WebpPreservesC2paChunkThatIsNotLastByDefault() {
        byte[] c2pa = CreateRiffChunk("C2PA", CreateManifestStore());
        byte[] keep = CreateRiffChunk("VP8 ", new byte[] { 1, 2 });
        byte[] webp = CreateWebp(c2pa, keep);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.False(result.WasChanged);
        Assert.False(result.Before.Evidence[0].IsStructurallyValid);
        Assert.Equal(webp, result.ToArray());
    }

    [Fact]
    public void AiMetadataRemovalDoesNotRemoveDisabledC2paCarriers() {
        var options = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = false,
            RemoveAiSourceMetadata = true
        };
        byte[] manifest = CreateManifestStore();
        byte[] webp = CreateWebp(CreateRiffChunk("VP8 ", new byte[] { 1, 2 }), CreateRiffChunk("C2PA", manifest));
        byte[] tiff = CreateLittleEndianTiff(manifest);
        byte[] svg = Encoding.UTF8.GetBytes(
            $"<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:c2pa=\"http://c2pa.org/manifest\"><metadata><c2pa:manifest>{Convert.ToBase64String(manifest)}</c2pa:manifest></metadata></svg>");

        OfficeProvenanceRemovalResult webpResult = OfficeProvenanceRemover.Remove(webp, "fixture.webp", options);
        OfficeProvenanceRemovalResult tiffResult = OfficeProvenanceRemover.Remove(tiff, "fixture.tif", options);
        OfficeProvenanceRemovalResult svgResult = OfficeProvenanceRemover.Remove(svg, "fixture.svg", options);

        Assert.Equal(webp, webpResult.ToArray());
        Assert.Equal(tiff, tiffResult.ToArray());
        Assert.Equal(svg, svgResult.ToArray());
        Assert.False(webpResult.WasChanged);
        Assert.False(tiffResult.WasChanged);
        Assert.False(svgResult.WasChanged);
    }

    [Fact]
    public void ProgressiveJpegAllowsMarkerSegmentsBetweenScans() {
        byte[] manifest = CreateManifestStore();
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegApp11(manifest, 0, manifest.Length, instance: 1, sequence: 1),
            CreateJpegSegment(0xDA, new byte[] { 1, 2 }),
            new byte[] { 0x11 },
            CreateJpegSegment(0xC4, new byte[] { 0, 1 }),
            CreateJpegSegment(0xDA, new byte[] { 3, 4 }),
            new byte[] { 0x22, 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "progressive.jpg");

        Assert.Single(result.Before.Evidence);
        Assert.True(result.Before.Evidence[0].IsStructurallyValid);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void GifRemovesOnlyTheExactC2paApplicationExtension() {
        byte[] manifest = CreateManifestStore();
        byte[] exact = CreateGifApplication("C2PA_GIF", new byte[] { 1, 0, 0 }, manifest);
        byte[] other = CreateGifApplication("C2PA_GIF", new byte[] { 1, 0, 1 }, Encoding.ASCII.GetBytes("keep"));
        byte[] gif = Join(Encoding.ASCII.GetBytes("GIF89a"), new byte[7], other, exact, new byte[] { 0x3B });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(gif, "fixture.gif");

        Assert.Equal(Join(Encoding.ASCII.GetBytes("GIF89a"), new byte[7], other, new byte[] { 0x3B }), result.ToArray());
        Assert.Single(result.Changes);
    }

    [Fact]
    public void GifAppliesTheManifestLimitOnlyToTheC2paApplicationExtension() {
        byte[] unrelated = CreateGifApplication("OTHERAPP", new byte[] { 1, 0, 0 }, new byte[512]);
        byte[] c2pa = CreateGifApplication("C2PA_GIF", new byte[] { 1, 0, 0 }, CreateManifestStore());
        byte[] gif = Join(Encoding.ASCII.GetBytes("GIF89a"), new byte[7], unrelated, c2pa, new byte[] { 0x3B });
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxAssetBytes = gif.Length + 1L;
        options.Limits.MaxManifestBytes = 64;

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(gif, "fixture.gif", options);

        Assert.Equal(Join(Encoding.ASCII.GetBytes("GIF89a"), new byte[7], unrelated, new byte[] { 0x3B }), result.ToArray());
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void TiffRemovesTheC2paTagWithoutMovingUnrelatedPayload() {
        byte[] manifest = CreateManifestStore();
        byte[] tiff = CreateLittleEndianTiff(manifest);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");
        byte[] output = result.ToArray();

        Assert.Equal((byte)0, output[8]);
        Assert.Equal((byte)0, output[9]);
        Assert.Equal(manifest, output.Skip(26).Take(manifest.Length).ToArray());
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void TiffCombinedRemovalRetainsTheCleanedXmpEntryCount() {
        byte[] originalXmp = CreateXmpPacket();
        byte[] tiff = CreateLittleEndianTiffWithC2paBeforeXmp(CreateManifestStore(), originalXmp);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");
        byte[] output = result.ToArray();

        Assert.True(result.WasChanged);
        Assert.Equal((ushort)1, BitConverter.ToUInt16(output, 8));
        Assert.Equal((ushort)700, BitConverter.ToUInt16(output, 10));
        uint cleanedLength = BitConverter.ToUInt32(output, 14);
        Assert.True(cleanedLength < originalXmp.Length);
        string cleaned = Encoding.UTF8.GetString(output, 38, checked((int)cleanedLength));
        Assert.DoesNotContain("trainedAlgorithmicMedia", cleaned, StringComparison.Ordinal);
        Assert.Contains("digitalCapture", cleaned, StringComparison.Ordinal);
    }

    [Fact]
    public void SvgRemovesOnlyNamespacedManifestElements() {
        string encoded = Convert.ToBase64String(CreateManifestStore());
        string svg = $"<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:c2pa=\"http://c2pa.org/manifest\"><metadata><keep>yes</keep><c2pa:manifest>{encoded}</c2pa:manifest></metadata></svg>";

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(Encoding.UTF8.GetBytes(svg), "fixture.svg");
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.True(result.WasReserialized);
        Assert.DoesNotContain("c2pa:manifest", output, StringComparison.Ordinal);
        Assert.Contains("<keep>yes</keep>", output, StringComparison.Ordinal);
    }

    [Fact]
    public void SvgXmpUsesTheAssetLimitInsteadOfTheManifestLimit() {
        string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\"><metadata>" +
            Encoding.UTF8.GetString(CreateXmpPacket()) +
            "</metadata><desc>" + new string('x', 256) + "</desc></svg>";
        byte[] data = Encoding.UTF8.GetBytes(svg);
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxAssetBytes = data.Length + 1L;
        options.Limits.MaxManifestBytes = 64;

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(data, "fixture.svg", options);

        Assert.True(result.Before.HasGenerativeAiDeclaration);
        Assert.False(result.After.HasGenerativeAiDeclaration);
        Assert.Contains(result.After.Evidence, item => item.DigitalSourceKind == OfficeProvenanceDigitalSourceKind.DigitalCapture);
        Assert.Contains(new string('x', 256), Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void ZipRemovesOnlyTheExactCaseSensitiveManifestEntry() {
        byte[] package = CreateZip(
            ("word/document.xml", Encoding.UTF8.GetBytes("<document/>")),
            ("META-INF/CONTENT_CREDENTIAL.C2PA", Encoding.ASCII.GetBytes("keep")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.docx");
        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);

        Assert.Null(archive.GetEntry("META-INF/content_credential.c2pa"));
        Assert.NotNull(archive.GetEntry("META-INF/CONTENT_CREDENTIAL.C2PA"));
        Assert.NotNull(archive.GetEntry("word/document.xml"));
        Assert.True(result.WasReserialized);
    }

    [Fact]
    public void ZipBlocksSignedPackageMutationByDefault() {
        byte[] package = CreateZip(
            ("_xmlsignatures/sig1.xml", Encoding.UTF8.GetBytes("<signature/>")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            OfficeProvenanceRemover.Remove(package, "signed.docx"));

        Assert.Contains("invalidate package signatures", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ZipBlocksProducerSpecificOdfSignatureMutationByDefault() {
        byte[] package = CreateZip(
            ("META-INF/customsignatures.xml", Encoding.UTF8.GetBytes("<signature/>")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            OfficeProvenanceRemover.Remove(package, "signed.odt"));

        Assert.Contains("invalidate package signatures", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ZipInspectsAndSanitizesSupportedEmbeddedImages() {
        byte[] header = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] image = Join(
            header,
            CreatePngChunk("IHDR", new byte[13]),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IEND", Array.Empty<byte>()));
        byte[] package = CreateZip(
            ("word/document.xml", Encoding.UTF8.GetBytes("<document/>")),
            ("word/media/image1.png", image));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(package, "fixture.docx");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.docx");
        byte[] cleanedImage = ReadZipEntry(result.ToArray(), "word/media/image1.png");

        Assert.Single(report.Evidence);
        Assert.StartsWith("ZIP/word/media/image1.png/PNG/caBX", report.Evidence[0].Location, StringComparison.Ordinal);
        Assert.Empty(OfficeProvenanceInspector.Inspect(cleanedImage, "image1.png").Evidence);
        Assert.Contains(result.Changes, item => item.Location.StartsWith("ZIP/word/media/image1.png/", StringComparison.Ordinal));
    }

    [Fact]
    public void ZipRemovalSkipsEmbeddedInspectionWhenDisabled() {
        byte[] image = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[13]),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IEND", Array.Empty<byte>()));
        byte[] package = CreateZip(
            ("META-INF/content_credential.c2pa", CreateManifestStore()),
            ("word/media/image1.png", image));
        var options = new OfficeProvenanceRemovalOptions { ProcessEmbeddedAssets = false };

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.docx", options);

        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.Equal(image, ReadZipEntry(result.ToArray(), "word/media/image1.png"));
    }

    [Fact]
    public void ZipPreservesMalformedEmbeddedSvgAndReportsADiagnostic() {
        byte[] malformedSvg = Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\"><broken></svg>");
        byte[] package = CreateZip(("word/media/image1.svg", malformedSvg));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(package, "fixture.docx");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.docx");

        Assert.Contains(report.Diagnostics, item => item.Contains("embedded asset was preserved", StringComparison.Ordinal));
        Assert.Equal(package, result.ToArray());
    }

    [Fact]
    public void ZipSignedPackageBlocksOnlyWhenARequestedMutationExists() {
        byte[] signedClean = CreateZip(
            ("_xmlsignatures/sig1.xml", Encoding.UTF8.GetBytes("<signature/>")),
            ("word/document.xml", Encoding.UTF8.GetBytes("<document/>")));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(signedClean, "clean-signed.docx");

        Assert.False(result.WasChanged);
        Assert.Equal(signedClean, result.ToArray());
    }

    [Fact]
    public void ZipEmbeddedAssetLimitIsEnforcedDuringInspection() {
        byte[] image = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IEND", Array.Empty<byte>()));
        byte[] package = CreateZip(("media/1.png", image), ("media/2.png", image));

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            package,
            "fixture.docx",
            new OfficeProvenanceOptions { MaxEmbeddedAssets = 1 }));
    }

    [Fact]
    public void ZipRewriteUsesTheExpandedContainerLimitForUnrelatedEntries() {
        byte[] unrelated = new byte[64 * 1024];
        byte[] package = CreateCompressedZip(
            ("large-unrelated.bin", unrelated),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxAssetBytes = package.Length + 1L;
        options.Limits.MaxManifestBytes = 64;
        options.Limits.MaxExpandedContainerBytes = 128 * 1024;

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.docx", options);

        Assert.Empty(result.After.Evidence);
        Assert.Equal(unrelated, ReadZipEntry(result.ToArray(), "large-unrelated.bin"));
    }

    [Fact]
    public void StructuredTextRemovesManifestBlockAndPreservesSurroundingText() {
        string block = "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n";
        byte[] text = Encoding.UTF8.GetBytes("before\n" + block + "after\n");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(text, "fixture.md");

        Assert.Equal("before\nafter\n", Encoding.UTF8.GetString(result.ToArray()));
        Assert.Single(result.Changes);
    }

    [Fact]
    public void StructuredTextExtensionWinsOverSvgTextAndHandlesLoneCarriageReturns() {
        string block = "-----BEGIN C2PA MANIFEST-----\r" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\r" +
            "-----END C2PA MANIFEST-----\r";
        byte[] text = Encoding.UTF8.GetBytes("literal <svg example\r" + block + "after\r");

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(text, "fixture.md");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(text, "fixture.md");

        Assert.Equal(OfficeProvenanceAssetFormat.StructuredText, report.Format);
        Assert.Equal("literal <svg example\rafter\r", Encoding.UTF8.GetString(result.ToArray()));
    }

    [Fact]
    public void StructuredTextIgnoresDelimiterWordsEmbeddedInProse() {
        byte[] text = Encoding.UTF8.GetBytes(
            "This mentions -----BEGIN C2PA MANIFEST----- but is not a carrier.\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "and -----END C2PA MANIFEST----- remains prose.\n");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(text, "fixture.md");

        Assert.False(result.WasChanged);
        Assert.Equal(text, result.ToArray());
    }

    [Fact]
    public void StructuredTextResynchronizesAtANewerStandaloneBeginDelimiter() {
        string validBlock = "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n";
        byte[] text = Encoding.UTF8.GetBytes("-----BEGIN C2PA MANIFEST-----\nstale\n" + validBlock + "after\n");

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(text, "fixture.md");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(text, "fixture.md");

        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.Equal("-----BEGIN C2PA MANIFEST-----\nstale\nafter\n", Encoding.UTF8.GetString(result.ToArray()));
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void UnstructuredTextRemovesOnlyACompleteC2paVariationSelectorWrapper() {
        byte[] wrapper = CreateTextWrapper(CreateManifestStore());
        byte[] normalVariationSelector = Encoding.UTF8.GetBytes("text️ stays ");
        byte[] text = Join(normalVariationSelector, wrapper, Encoding.UTF8.GetBytes("tail"));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(text, "fixture.txt");

        Assert.Equal(Join(normalVariationSelector, Encoding.UTF8.GetBytes("tail")), result.ToArray());
        Assert.Single(result.Changes);
    }

    [Fact]
    public void InspectionEnforcesAssetAndCarrierBounds() {
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            new byte[9],
            "fixture.bin",
            new OfficeProvenanceOptions { MaxAssetBytes = 8, MaxManifestBytes = 8 }));

        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IEND", Array.Empty<byte>()));
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            png,
            "fixture.png",
            new OfficeProvenanceOptions { MaxCarriers = 1 }));
    }

    [Fact]
    public void JpegXmpRemovesOnlyAiDigitalSourceTypeDeclarations() {
        byte[] xmp = CreateXmpPacket();
        byte[] xmpHeader = Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0");
        byte[] jpeg = Join(new byte[] { 0xFF, 0xD8 }, CreateJpegSegment(0xE1, Join(xmpHeader, xmp)), new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.True(result.Before.HasGenerativeAiDeclaration);
        Assert.False(result.After.HasGenerativeAiDeclaration);
        Assert.Contains(result.After.Evidence, item => item.DigitalSourceKind == OfficeProvenanceDigitalSourceKind.DigitalCapture);
        Assert.Contains(result.Changes, item => item.Carrier == OfficeProvenanceCarrierKind.IptcDigitalSourceType);
        Assert.True(result.WasReserialized);
        Assert.Single(result.After.Evidence);
    }

    [Fact]
    public void PngXmpRemovalRewritesAValidItextChunk() {
        byte[] header = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] prefix = Join(Encoding.ASCII.GetBytes("XML:com.adobe.xmp"), new byte[] { 0, 0, 0, 0, 0 });
        byte[] png = Join(header, CreatePngChunk("iTXt", Join(prefix, CreateXmpPacket())), CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.True(result.Before.HasGenerativeAiDeclaration);
        Assert.False(result.After.HasGenerativeAiDeclaration);
        Assert.True(result.WasReserialized);
        Assert.Contains(result.After.Evidence, item => item.DigitalSourceKind == OfficeProvenanceDigitalSourceKind.DigitalCapture);
    }

    [Fact]
    public void PngXmpWithInvalidCrcIsReportedAndPreservedByDefault() {
        byte[] header = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] prefix = Join(Encoding.ASCII.GetBytes("XML:com.adobe.xmp"), new byte[] { 0, 0, 0, 0, 0 });
        byte[] xmp = CreatePngChunk("iTXt", Join(prefix, CreateXmpPacket()));
        xmp[xmp.Length - 1] ^= 0x01;
        byte[] png = Join(header, xmp, CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.Contains(result.Before.Evidence, item => !item.IsStructurallyValid);
        Assert.Equal(png, result.ToArray());
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void WebpXmpRemovalPreservesNonAiDigitalSourceType() {
        byte[] webp = CreateWebp(CreateRiffChunk("VP8 ", new byte[] { 1, 2 }), CreateRiffChunk("XMP ", CreateXmpPacket()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.True(result.Before.HasGenerativeAiDeclaration);
        Assert.False(result.After.HasGenerativeAiDeclaration);
        Assert.True(result.WasReserialized);
        Assert.Contains(result.After.Evidence, item => item.DigitalSourceKind == OfficeProvenanceDigitalSourceKind.DigitalCapture);
    }

    private static byte[] CreateManifestStore(int length = 38) {
        if (length < 38) throw new ArgumentOutOfRangeException(nameof(length));
        byte[] data = new byte[length];
        WriteBigEndian(data, 0, length);
        Encoding.ASCII.GetBytes("jumb").CopyTo(data, 4);
        WriteBigEndian(data, 8, 30);
        Encoding.ASCII.GetBytes("jumd").CopyTo(data, 12);
        new byte[] { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 }.CopyTo(data, 16);
        data[32] = 0x02;
        Encoding.ASCII.GetBytes("c2pa").CopyTo(data, 33);
        return data;
    }

    private static byte[] CreateExtendedManifestStore(int length) {
        if (length < 46) throw new ArgumentOutOfRangeException(nameof(length));
        byte[] data = new byte[length];
        WriteBigEndian(data, 0, 1);
        Encoding.ASCII.GetBytes("jumb").CopyTo(data, 4);
        WriteBigEndian64(data, 8, (ulong)length);
        WriteBigEndian(data, 16, 30);
        Encoding.ASCII.GetBytes("jumd").CopyTo(data, 20);
        new byte[] { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 }.CopyTo(data, 24);
        data[40] = 0x02;
        Encoding.ASCII.GetBytes("c2pa").CopyTo(data, 41);
        return data;
    }

    private static byte[] CreateXmpPacket() => Encoding.UTF8.GetBytes(
        "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
        "<rdf:Description xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" +
        "<iptc:DigitalSourceType rdf:resource=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
        "<iptc:DigitalSourceType rdf:resource=\"http://cv.iptc.org/newscodes/digitalsourcetype/digitalCapture\"/>" +
        "</rdf:Description></rdf:RDF></x:xmpmeta>");

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

    private static byte[] CreateJpegSegment(byte marker, byte[] payload) {
        byte[] segment = new byte[payload.Length + 4];
        segment[0] = 0xFF;
        segment[1] = marker;
        int length = payload.Length + 2;
        segment[2] = (byte)(length >> 8);
        segment[3] = (byte)length;
        payload.CopyTo(segment, 4);
        return segment;
    }

    private static byte[] CreatePngChunk(string type, byte[] payload) {
        byte[] chunk = new byte[payload.Length + 12];
        WriteBigEndian(chunk, 0, payload.Length);
        Encoding.ASCII.GetBytes(type).CopyTo(chunk, 4);
        payload.CopyTo(chunk, 8);
        WriteBigEndian(chunk, chunk.Length - 4, unchecked((int)ComputePngCrc(chunk, 4, payload.Length + 4)));
        return chunk;
    }

    private static uint ComputePngCrc(byte[] data, int offset, int count) {
        uint crc = 0xFFFFFFFF;
        for (int index = offset; index < offset + count; index++) {
            crc ^= data[index];
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
        }
        return crc ^ 0xFFFFFFFF;
    }

    private static byte[] CreateRiffChunk(string type, byte[] payload) {
        byte[] chunk = new byte[8 + payload.Length + (payload.Length & 1)];
        Encoding.ASCII.GetBytes(type).CopyTo(chunk, 0);
        BitConverter.GetBytes(payload.Length).CopyTo(chunk, 4);
        payload.CopyTo(chunk, 8);
        return chunk;
    }

    private static byte[] CreateWebp(params byte[][] chunks) {
        byte[] body = Join(chunks);
        byte[] result = new byte[12 + body.Length];
        Encoding.ASCII.GetBytes("RIFF").CopyTo(result, 0);
        BitConverter.GetBytes(result.Length - 8).CopyTo(result, 4);
        Encoding.ASCII.GetBytes("WEBP").CopyTo(result, 8);
        body.CopyTo(result, 12);
        return result;
    }

    private static byte[] CreateGifApplication(string identifier, byte[] authenticationCode, byte[] payload) {
        byte[] header = Join(new byte[] { 0x21, 0xFF, 0x0B }, Encoding.ASCII.GetBytes(identifier), authenticationCode);
        using var output = new MemoryStream();
        output.Write(header, 0, header.Length);
        int offset = 0;
        while (offset < payload.Length) {
            int count = Math.Min(255, payload.Length - offset);
            output.WriteByte((byte)count);
            output.Write(payload, offset, count);
            offset += count;
        }
        output.WriteByte(0);
        return output.ToArray();
    }

    private static byte[] CreateLittleEndianTiff(byte[] manifest) {
        byte[] result = new byte[26 + manifest.Length];
        result[0] = result[1] = (byte)'I';
        result[2] = 42;
        result[4] = 8;
        result[8] = 1;
        result[10] = 0x41;
        result[11] = 0xCD;
        result[12] = 7;
        BitConverter.GetBytes(manifest.Length).CopyTo(result, 14);
        BitConverter.GetBytes(26).CopyTo(result, 18);
        manifest.CopyTo(result, 26);
        return result;
    }

    private static byte[] CreateLittleEndianTiffWithC2paBeforeXmp(byte[] manifest, byte[] xmp) {
        const int payloadOffset = 38;
        byte[] result = new byte[payloadOffset + manifest.Length + xmp.Length];
        result[0] = result[1] = (byte)'I';
        result[2] = 42;
        result[4] = 8;
        result[8] = 2;
        WriteLittleEndianEntry(result, 10, 0xCD41, 7, manifest.Length, payloadOffset);
        WriteLittleEndianEntry(result, 22, 700, 1, xmp.Length, payloadOffset + manifest.Length);
        Buffer.BlockCopy(manifest, 0, result, payloadOffset, manifest.Length);
        Buffer.BlockCopy(xmp, 0, result, payloadOffset + manifest.Length, xmp.Length);
        return result;
    }

    private static void WriteLittleEndianEntry(byte[] data, int offset, ushort tag, ushort type, int count, int valueOffset) {
        BitConverter.GetBytes(tag).CopyTo(data, offset);
        BitConverter.GetBytes(type).CopyTo(data, offset + 2);
        BitConverter.GetBytes(count).CopyTo(data, offset + 4);
        BitConverter.GetBytes(valueOffset).CopyTo(data, offset + 8);
    }

    private static byte[] CreateZip(params (string Name, byte[] Data)[] entries) {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach ((string name, byte[] data) in entries) {
                ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
                using Stream target = entry.Open();
                target.Write(data, 0, data.Length);
            }
        }
        return stream.ToArray();
    }

    private static byte[] CreateCompressedZip(params (string Name, byte[] Data)[] entries) {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach ((string name, byte[] data) in entries) {
                ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
                using Stream target = entry.Open();
                target.Write(data, 0, data.Length);
            }
        }
        return stream.ToArray();
    }

    private static byte[] ReadZipEntry(byte[] package, string name) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        ZipArchiveEntry entry = archive.GetEntry(name) ?? throw new InvalidDataException("ZIP fixture entry was not found.");
        using Stream source = entry.Open();
        using var output = new MemoryStream();
        source.CopyTo(output);
        return output.ToArray();
    }

    private static byte[] CreateTextWrapper(byte[] manifest) {
        byte[] header = Join(Encoding.ASCII.GetBytes("C2PATXT\0"), new byte[] { 1 }, BigEndian(manifest.Length), manifest);
        var builder = new StringBuilder("\uFEFF");
        foreach (byte value in header) builder.Append(char.ConvertFromUtf32(value < 16 ? 0xFE00 + value : 0xE0100 + value - 16));
        return Encoding.UTF8.GetBytes(builder.ToString());
    }

    private static byte[] BigEndian(int value) {
        byte[] result = new byte[4];
        WriteBigEndian(result, 0, value);
        return result;
    }

    private static void WriteBigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static void WriteBigEndian64(byte[] data, int offset, ulong value) {
        data[offset] = (byte)(value >> 56);
        data[offset + 1] = (byte)(value >> 48);
        data[offset + 2] = (byte)(value >> 40);
        data[offset + 3] = (byte)(value >> 32);
        data[offset + 4] = (byte)(value >> 24);
        data[offset + 5] = (byte)(value >> 16);
        data[offset + 6] = (byte)(value >> 8);
        data[offset + 7] = (byte)value;
    }

    private static byte[] Join(params byte[][] parts) {
        int length = parts.Sum(part => part.Length);
        byte[] result = new byte[length];
        int offset = 0;
        foreach (byte[] part in parts) {
            Buffer.BlockCopy(part, 0, result, offset, part.Length);
            offset += part.Length;
        }
        return result;
    }
}

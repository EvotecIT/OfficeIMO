using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData(1, 0x40, "", 0, 255)]
    [InlineData(1, 0x40, " /Decode [1 0]", 255, 0)]
    [InlineData(2, 0x30, "", 0, 255)]
    [InlineData(4, 0x0F, "", 0, 255)]
    public void RenderPage_ProjectsPackedGrayThroughDecodePipeline(int bits, int packed, string extra, byte first, byte second) {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(new[] { (byte)packed },
            colorSpace: "/DeviceGray", bitsPerComponent: bits, imageWidth: 2,
            extraImageEntries: extra, imageFilterEntry: "");
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingImage image = Assert.Single(drawing.Images);
        Assert.True(OfficeRasterImageDecoder.TryDecode(image.Bytes, out OfficeRasterImage? decoded));
        Assert.Equal(OfficeColor.FromRgb(first, first, first), decoded!.GetPixel(0, 0));
        Assert.Equal(OfficeColor.FromRgb(second, second, second), decoded.GetPixel(1, 0));
        Assert.DoesNotContain(PdfPageImageRenderer.RenderPages(pdf).Single().CapabilityDiagnostics,
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void RenderPage_PreservesPackedGrayColorKeyMask() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(new byte[] { 0x40 },
            colorSpace: "/DeviceGray", bitsPerComponent: 1, imageWidth: 2,
            extraImageEntries: " /Mask [1 1]", imageFilterEntry: "");
        OfficeDrawingImage image = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Images);
        Assert.True(OfficeRasterImageDecoder.TryDecode(image.Bytes, out OfficeRasterImage? decoded));
        Assert.Equal(OfficeColor.Black, decoded!.GetPixel(0, 0));
        Assert.Equal(0, decoded.GetPixel(1, 0).A);
    }

    [Fact]
    public async Task Ocr_RendersCcittAndRejectsMalformedScanBeforeProviderCall() {
        byte[] valid = BuildSingleStreamPdfWithBinaryImageXObject(PdfFaxDecodeTests.Pack("001 00110101 000101"),
            colorSpace: "/DeviceGray", bitsPerComponent: 1, imageWidth: 8,
            imageFilterEntry: "/Filter /CCITTFaxDecode /DecodeParms << /K -1 /Columns 8 /Rows 1 /EndOfBlock false >>");
        int calls = 0;
        var engine = new DelegateOcrEngine("scan-proof", (request, token) => {
            calls++;
            Assert.True(OfficeRasterImageDecoder.TryDecode(request.Payload, out OfficeRasterImage? raster));
            Assert.Equal(OfficeColor.Black, raster!.GetPixel(50, 110));
            return Task.FromResult(new OcrResult());
        });
        await PdfDocument.Load(valid).ReadWithOcrAsync(engine, new PdfOcrMergeOptions { Dpi = 72 });
        Assert.Equal(1, calls);
        byte[] invalid = BuildSingleStreamPdfWithBinaryImageXObject(new byte[] { 0 },
            colorSpace: "/DeviceGray", bitsPerComponent: 1, imageWidth: 8,
            imageFilterEntry: "/Filter /CCITTFaxDecode /DecodeParms << /K -1 /Columns 8 /Rows 1 /EndOfBlock false >>");
        PdfPageRenderResult failed = Assert.Single(PdfPageImageRenderer.RenderPages(invalid));
        Assert.False(failed.Succeeded);
        await Assert.ThrowsAsync<NotSupportedException>(() => PdfDocument.Load(invalid).ReadWithOcrAsync(engine));
        Assert.Equal(1, calls);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Ocr_UsesJpeg2000CodecAndDoesNotSendBlankPageWhenCodecIsMissing(bool rawCodestream) {
        byte[] payload = ReadScanJpx("rgb");
        if (rawCodestream) payload = payload.Skip(FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51)).ToArray();
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(payload,
            colorSpace: "/DeviceRGB", imageWidth: 1, imageFilterEntry: "/Filter /JPXDecode");
        var codec = new ScanJpxCodec(payload);
        int calls = 0;
        var engine = new DelegateOcrEngine("codec-proof", (request, token) => {
            calls++;
            Assert.True(OfficeRasterImageDecoder.TryDecode(request.Payload, out OfficeRasterImage? raster));
            Assert.Equal(OfficeColor.Red, raster!.GetPixel(50, 110));
            return Task.FromResult(new OcrResult());
        });
        await Assert.ThrowsAsync<NotSupportedException>(() => PdfDocument.Load(pdf).ReadWithOcrAsync(engine));
        Assert.Equal(0, calls);
        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { Dpi = 72, ImageCodec = codec });
        Assert.Equal(1, calls);
        Assert.Equal(1, codec.Calls);
        Assert.Contains(result.Pages.Single().Diagnostics, diagnostic => diagnostic.StartsWith(PdfRenderCapabilities.OptionalImageCodecId));
    }

    [Theory]
    [InlineData("rgba", "")]
    [InlineData("rgba", " /SMaskInData 0")]
    [InlineData("rgba", " /SMaskInData 1")]
    [InlineData("rgba", " /SMaskInData 2")]
    [InlineData("rgb", " /SMaskInData 1")]
    [InlineData("rgb", " /SMaskInData 2")]
    public async Task Ocr_RejectsJpeg2000OpacityBeforeRenderingOrCallingProvider(string mode, string mask) {
        byte[] payload = ReadScanJpx(mode);
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(payload,
            colorSpace: "/DeviceRGB", imageWidth: 1, imageFilterEntry: "/Filter /JPXDecode",
            extraImageEntries: mask);
        var codec = new ScanJpxCodec(payload);
        Assert.False(Assert.Single(PdfPageImageRenderer.RenderPages(pdf,
            options: new PdfPageRenderOptions { ImageCodec = codec })).Succeeded);
        int calls = 0;
        var engine = new DelegateOcrEngine("alpha-proof", (request, token) => {
            calls++;
            return Task.FromResult(new OcrResult());
        });
        await Assert.ThrowsAsync<NotSupportedException>(() => PdfDocument.Load(pdf).ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { ImageCodec = codec }));
        Assert.Equal(0, calls);
        Assert.Equal(0, codec.Calls);
    }

    [Fact]
    public void Jpeg2000Header_RejectsTruncatedAndOversizedBoxesAndRecognizesOpaqueCodestream() {
        byte[] payload = ReadScanJpx("rgb");
        Assert.True(OfficeJpeg2000Header.TryGetOpaqueComponents(payload, out int components));
        Assert.Equal(3, components);
        for (int length = 0; length < payload.Length; length++) {
            Assert.False(OfficeJpeg2000Header.TryGetOpaqueComponents(payload.Take(length).ToArray(), out _));
        }
        byte[] oversized = (byte[])payload.Clone();
        for (int i = 12; i < 16; i++) oversized[i] = 255;
        Assert.False(OfficeJpeg2000Header.TryGetOpaqueComponents(oversized, out _));
        int marker = Enumerable.Range(0, payload.Length - 3).Single(index =>
            payload[index] == 255 && payload[index + 1] == 79 && payload[index + 2] == 255 && payload[index + 3] == 81);
        Assert.True(OfficeJpeg2000Header.TryGetOpaqueComponents(payload.Skip(marker).ToArray(), out components));
        Assert.Equal(3, components);
        Assert.False(OfficeJpeg2000Header.TryGetOpaqueComponents(ReadScanJpx("rgba"), out _));
    }

    [Fact]
    public void ImageValidationRejectsJpeg2000CodestreamTruncatedAfterItsSizeHeader() {
        byte[] payload = ReadScanJpx("rgb");
        int marker = Enumerable.Range(0, payload.Length - 3).Single(index =>
            payload[index] == 255 && payload[index + 1] == 79 && payload[index + 2] == 255 && payload[index + 3] == 81);
        byte[] rawCodestream = payload.Skip(marker).ToArray();
        int sizeSegmentLength = (rawCodestream[4] << 8) | rawCodestream[5];
        byte[] truncated = rawCodestream.Take(4 + sizeSegmentLength).ToArray();

        Assert.True(OfficeImageReader.TryIdentifyByContent(truncated, "scan.j2k", out OfficeImageInfo identified));
        Assert.Equal(OfficeImageFormat.Jpeg2000Codestream, identified.Format);
        Assert.Equal("image/j2c", identified.MimeType);
        Assert.Equal(".j2c", OfficeImageInfo.GetDefaultExtension(identified.Format));
        Assert.False(OfficeImageReader.TryValidateContent(truncated, "scan.j2k", out _));
        Assert.True(OfficeImageReader.TryValidateContent(rawCodestream, "scan.j2k", out OfficeImageInfo validated));
        Assert.Equal(OfficeImageFormat.Jpeg2000Codestream, validated.Format);
        Assert.Equal("image/j2c", validated.MimeType);

        Assert.True(OfficeImageReader.TryValidateContent(payload, "scan.jp2", out OfficeImageInfo container));
        Assert.Equal(OfficeImageFormat.Jpeg2000, container.Format);
        Assert.Equal("image/jp2", container.MimeType);
    }

    [Theory]
    [InlineData(false, 1)]
    [InlineData(false, 65535)]
    [InlineData(true, 1)]
    [InlineData(true, 65535)]
    public void ImageValidationRejectsNonBaselineJpeg2000Capabilities(bool rawCodestream, int capabilities) {
        byte[] container = ReadScanJpx("rgb");
        int marker = Enumerable.Range(0, container.Length - 3).Single(index =>
            container[index] == 255 && container[index + 1] == 79 &&
            container[index + 2] == 255 && container[index + 3] == 81);
        byte[] payload = rawCodestream ? container.Skip(marker).ToArray() : container;
        int sizeMarker = rawCodestream ? 0 : marker;
        Assert.Equal(0, (payload[sizeMarker + 6] << 8) | payload[sizeMarker + 7]);
        payload[sizeMarker + 6] = (byte)(capabilities >> 8);
        payload[sizeMarker + 7] = (byte)capabilities;

        Assert.False(OfficeImageReader.TryValidateContent(payload,
            rawCodestream ? "scan.j2c" : "scan.jp2", out _));
    }

    [Fact]
    public void ImageValidationRejectsJp2WithoutMandatoryFileTypeBox() {
        byte[] payload = ReadScanJpx("rgb");
        int fileTypeLength = ReadUInt32BigEndian(payload, 12);
        byte[] missingFileType = payload.Take(12).Concat(payload.Skip(12 + fileTypeLength)).ToArray();

        Assert.False(OfficeImageReader.TryValidateContent(missingFileType, "scan.jp2", out _));
    }

    [Theory]
    [InlineData(20)] // Brand must declare the baseline JP2 format.
    [InlineData(28)] // Compatibility list must include the baseline JP2 format.
    public void ImageValidationRejectsJp2WithMalformedFileTypeBox(int fieldOffset) {
        byte[] payload = ReadScanJpx("rgb");
        payload[fieldOffset] = (byte)'x';

        Assert.False(OfficeImageReader.TryValidateContent(payload, "scan.jp2", out _));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(65535)]
    public void ImageValidationRejectsJp2WithNonbaselineFileTypeMinorVersion(int minorVersion) {
        byte[] payload = ReadScanJpx("rgb");
        payload[24] = (byte)(minorVersion >> 24);
        payload[25] = (byte)(minorVersion >> 16);
        payload[26] = (byte)(minorVersion >> 8);
        payload[27] = (byte)minorVersion;

        Assert.False(OfficeImageReader.TryValidateContent(payload, "scan.jp2", out _));
    }

    [Fact]
    public void ImageValidationRejectsJp2WithUndersizedUuidBox() {
        byte[] payload = ReadScanJpx("rgb");
        int codestreamType = FindMarker(payload, (byte)'j', (byte)'p', (byte)'2', (byte)'c');
        int codestreamBox = codestreamType - 4;
        byte[] malformed = payload.Take(codestreamBox)
            .Concat(CreateJp2Box("uuid", new byte[15]))
            .Concat(payload.Skip(codestreamBox))
            .ToArray();

        Assert.False(OfficeJpeg2000Header.TryValidateOpaquePayload(malformed, out _, out _, out _));
        Assert.False(OfficeImageReader.TryValidateContent(malformed, "scan.jp2", out _));
    }

    [Theory]
    [InlineData(10, 38)] // Component precision is limited to 38 bits (encoded as precision minus one).
    [InlineData(11, 6)] // Compression type must be JPEG 2000 (7).
    [InlineData(12, 2)] // Unknown-colourspace flag is boolean.
    [InlineData(13, 1)] // Intellectual-property metadata is outside the accepted opaque subset.
    [InlineData(13, 2)] // Intellectual-property flag must still be Boolean.
    [InlineData(10, 255)] // Variable component depths require a matching bpcc box.
    public void ImageValidationRejectsJp2WithInvalidImageHeaderField(int fieldOffset, int invalidValue) {
        byte[] payload = ReadScanJpx("rgb");
        int imageHeaderType = FindMarker(payload, 0x69, 0x68, 0x64, 0x72);
        payload[imageHeaderType + 4 + fieldOffset] = (byte)invalidValue;

        Assert.False(OfficeImageReader.TryValidateContent(payload, "scan.jp2", out _));
    }

    [Theory]
    [InlineData(11)] // Valid unsigned 12-bit declaration does not match the codestream's unsigned 8-bit Ssiz.
    [InlineData(135)] // Valid signed 8-bit declaration does not match the codestream's unsigned 8-bit Ssiz.
    public void ImageValidationRejectsJp2WhenImageHeaderPrecisionDiffersFromCodestream(int headerPrecision) {
        byte[] payload = ReadScanJpx("rgb");
        int imageHeaderType = FindMarker(payload, 0x69, 0x68, 0x64, 0x72);
        payload[imageHeaderType + 14] = (byte)headerPrecision;

        Assert.False(OfficeImageReader.TryValidateContent(payload, "scan.jp2", out _));
    }

    [Theory]
    [InlineData(1)] // PRECEDENCE must be zero in the baseline JP2 subset.
    [InlineData(2)] // APPROX must be zero for an exact baseline colourspace.
    public void ImageValidationRejectsJp2WithNonzeroColourSpecificationFlag(int fieldOffset) {
        byte[] payload = ReadScanJpx("rgb");
        int colorSpecificationType = FindMarker(payload, 0x63, 0x6F, 0x6C, 0x72);
        int colorSpecificationContent = colorSpecificationType + 4;
        Assert.Equal(0, payload[colorSpecificationContent + fieldOffset]);
        payload[colorSpecificationContent + fieldOffset] = 1;

        Assert.False(OfficeJpeg2000Header.TryGetOpaqueDimensions(payload, out _, out _, out _));
        Assert.False(OfficeJpeg2000Header.TryValidateOpaquePayload(payload, out _, out _, out _));
        Assert.False(OfficeImageReader.TryValidateContent(payload, "scan.jp2", out _));
    }

    [Theory]
    [InlineData("resc")]
    [InlineData("resd")]
    [InlineData("both")]
    public void ImageValidationAcceptsCompleteJp2ResolutionSuperbox(string resolutionType) {
        byte[] resolution = { 0, 1, 0, 1, 0, 1, 0, 1, 0, 0 };
        byte[] children = resolutionType == "both"
            ? CreateJp2Box("resc", resolution).Concat(CreateJp2Box("resd", resolution)).ToArray()
            : CreateJp2Box(resolutionType, resolution);
        byte[] payload = InsertJp2ResolutionBox(ReadScanJpx("rgb"),
            CreateJp2Box("res ", children));

        Assert.True(OfficeJpeg2000Header.TryValidateOpaquePayload(payload, out _, out _, out _));
        Assert.True(OfficeImageReader.TryValidateContent(payload, "scan.jp2", out _));
    }

    [Fact]
    public void ImageValidationRejectsMalformedJp2ResolutionSuperboxes() {
        byte[] validResolution = { 0, 1, 0, 1, 0, 1, 0, 1, 0, 0 };
        byte[] zeroDenominator = (byte[])validResolution.Clone();
        zeroDenominator[3] = 0;
        byte[] validChild = CreateJp2Box("resc", validResolution);
        byte[] validBox = CreateJp2Box("res ", validChild);
        byte[][] malformed = {
            CreateJp2Box("res ", Array.Empty<byte>()),
            CreateJp2Box("res ", CreateJp2Box("resc", validResolution.Take(9).ToArray())),
            CreateJp2Box("res ", CreateJp2Box("resc", new byte[10])),
            CreateJp2Box("res ", CreateJp2Box("resc", zeroDenominator)),
            CreateJp2Box("res ", CreateJp2Box("junk", validResolution)),
            CreateJp2Box("res ", validChild.Concat(validChild).ToArray()),
            validBox.Concat(validBox).ToArray()
        };
        foreach (byte[] resolutionBox in malformed) {
            byte[] payload = InsertJp2ResolutionBox(ReadScanJpx("rgb"), resolutionBox);
            Assert.False(OfficeJpeg2000Header.TryValidateOpaquePayload(payload, out _, out _, out _));
            Assert.False(OfficeImageReader.TryValidateContent(payload, "scan.jp2", out _));
        }
    }

    [Theory]
    [InlineData(0x52)] // COD: coding-style default.
    [InlineData(0x5C)] // QCD: quantization default.
    public void ImageValidationRejectsJpeg2000CodestreamWithoutMandatoryMainHeaderMarker(int markerCode) {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] rawCodestream = payload.Skip(codestream).ToArray();
        int marker = FindMarker(rawCodestream, 0xFF, (byte)markerCode);
        int markerLength = (rawCodestream[marker + 2] << 8) | rawCodestream[marker + 3];
        byte[] malformed = rawCodestream.Take(marker)
            .Concat(rawCodestream.Skip(marker + 2 + markerLength))
            .ToArray();

        Assert.True(OfficeImageReader.TryIdentifyByContent(malformed, "scan.j2c", out _));
        Assert.False(OfficeImageReader.TryValidateContent(malformed, "scan.j2c", out _));
    }

    [Theory]
    [InlineData(0x52)] // COD cannot be an empty marker segment.
    [InlineData(0x5C)] // QCD cannot be an empty marker segment.
    public void ImageValidationRejectsJpeg2000CodestreamWithEmptyMandatoryMainHeaderSegment(int markerCode) {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] rawCodestream = payload.Skip(codestream).ToArray();
        int marker = FindMarker(rawCodestream, 0xFF, (byte)markerCode);
        int markerLength = (rawCodestream[marker + 2] << 8) | rawCodestream[marker + 3];
        byte[] malformed = rawCodestream.Take(marker)
            .Concat(new byte[] { 0xFF, (byte)markerCode, 0x00, 0x02 })
            .Concat(rawCodestream.Skip(marker + 2 + markerLength))
            .ToArray();

        Assert.True(OfficeImageReader.TryIdentifyByContent(malformed, "scan.j2c", out _));
        Assert.False(OfficeImageReader.TryValidateContent(malformed, "scan.j2c", out _));
    }

    [Theory]
    [InlineData(0x52)] // An optional tile COD override still requires a complete COD body.
    [InlineData(0x5C)] // An optional tile QCD override still requires a complete QCD body.
    public void ImageValidationRejectsJpeg2000CodestreamWithEmptyTileHeaderOverrideSegment(int markerCode) {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] rawCodestream = payload.Skip(codestream).ToArray();
        int tilePart = FindMarker(rawCodestream, 0xFF, 0x90);
        int startOfData = FindMarker(rawCodestream, 0xFF, 0x93);
        uint tilePartLength = (uint)ReadUInt32BigEndian(rawCodestream, tilePart + 6);
        byte[] malformed = rawCodestream.Take(startOfData)
            .Concat(new byte[] { 0xFF, (byte)markerCode, 0x00, 0x02 })
            .Concat(rawCodestream.Skip(startOfData))
            .ToArray();
        WriteJpxUInt32(malformed, tilePart + 6, tilePartLength + 4U);

        Assert.True(OfficeImageReader.TryIdentifyByContent(malformed, "scan.j2c", out _));
        Assert.False(OfficeImageReader.TryValidateContent(malformed, "scan.j2c", out _));
    }

    [Theory]
    [InlineData(0x53)] // COC changes component coding style.
    [InlineData(0x5D)] // QCC changes component quantization.
    [InlineData(0x5E)] // RGN changes region-of-interest decoding.
    [InlineData(0x5F)] // POC changes progression order.
    public void ImageValidationRejectsUnvalidatedJpeg2000MainAndTileMarkers(int markerCode) {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] raw = payload.Skip(codestream).ToArray();
        int tilePart = FindMarker(raw, 0xFF, 0x90);
        byte[] segment = { 0xFF, (byte)markerCode, 0x00, 0x02 };
        byte[] malformedMain = raw.Take(tilePart).Concat(segment).Concat(raw.Skip(tilePart)).ToArray();
        Assert.False(OfficeImageReader.TryValidateContent(malformedMain, "scan.j2c", out _));

        int startOfData = FindMarker(raw, 0xFF, 0x93);
        uint tilePartLength = (uint)ReadUInt32BigEndian(raw, tilePart + 6);
        byte[] malformedTile = raw.Take(startOfData).Concat(segment).Concat(raw.Skip(startOfData)).ToArray();
        WriteJpxUInt32(malformedTile, tilePart + 6, tilePartLength + (uint)segment.Length);
        Assert.False(OfficeImageReader.TryValidateContent(malformedTile, "scan.j2c", out _));
    }

    [Fact]
    public void ImageValidationRejectsMalformedJpeg2000CommentMarker() {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] raw = payload.Skip(codestream).ToArray();
        int tilePart = FindMarker(raw, 0xFF, 0x90);
        byte[] malformed = raw.Take(tilePart)
            .Concat(new byte[] { 0xFF, 0x64, 0x00, 0x04, 0x00, 0x02 })
            .Concat(raw.Skip(tilePart)).ToArray();

        Assert.False(OfficeImageReader.TryValidateContent(malformed, "scan.j2c", out _));
    }

    [Fact]
    public void ImageValidationAcceptsValidJpeg2000CommentMarkers() {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] raw = payload.Skip(codestream).ToArray();
        int tilePart = FindMarker(raw, 0xFF, 0x90);
        byte[] comment = { 0xFF, 0x64, 0x00, 0x04, 0x00, 0x01 };
        byte[] withMainComment = raw.Take(tilePart).Concat(comment).Concat(raw.Skip(tilePart)).ToArray();
        Assert.True(OfficeImageReader.TryValidateContent(withMainComment, "scan.j2c", out _));

        int startOfData = FindMarker(raw, 0xFF, 0x93);
        uint tilePartLength = (uint)ReadUInt32BigEndian(raw, tilePart + 6);
        byte[] withTileComment = raw.Take(startOfData).Concat(comment).Concat(raw.Skip(startOfData)).ToArray();
        WriteJpxUInt32(withTileComment, tilePart + 6, tilePartLength + (uint)comment.Length);
        Assert.True(OfficeImageReader.TryValidateContent(withTileComment, "scan.j2c", out _));
    }

    [Theory]
    [InlineData(0, 0x08)] // Scod reserved bits must be zero.
    [InlineData(1, 5)] // Progression order is limited to the five Part 1 orders.
    [InlineData(3, 0)] // At least one quality layer is required.
    [InlineData(4, 2)] // Multiple-component transform is boolean.
    [InlineData(5, 33)] // Decomposition levels are limited to 32.
    [InlineData(6, 9)] // Code-block width exponent exceeds the Part 1 limit.
    [InlineData(8, 0x40)] // Reserved code-block style bits must be zero.
    [InlineData(9, 2)] // Only the reversible and irreversible Part 1 transforms are valid.
    public void ImageValidationRejectsJpeg2000CodestreamWithInvalidCodingStyleDefaultField(
        int fieldOffset,
        int invalidValue) {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] rawCodestream = payload.Skip(codestream).ToArray();
        int codingStyle = FindMarker(rawCodestream, 0xFF, 0x52);
        rawCodestream[codingStyle + 4 + fieldOffset] = (byte)invalidValue;

        Assert.False(OfficeImageReader.TryValidateContent(rawCodestream, "scan.j2c", out _));
    }

    [Fact]
    public void ImageValidationRejectsComponentTransformInOneComponentCodestream() {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        var bytes = payload.Skip(codestream).ToList();
        Assert.Equal(47, (bytes[4] << 8) | bytes[5]);
        bytes.RemoveRange(45, 6); // Keep one SIZ component specification.
        bytes[5] = 41;
        bytes[40] = 0;
        bytes[41] = 1;
        byte[] oneComponent = bytes.ToArray();
        int codingStyle = FindMarker(oneComponent, 0xFF, 0x52);
        oneComponent[codingStyle + 8] = 0;
        Assert.True(OfficeImageReader.TryValidateContent(oneComponent, "scan.j2c", out _));
        oneComponent[codingStyle + 8] = 1;
        Assert.False(OfficeImageReader.TryValidateContent(oneComponent, "scan.j2c", out _));
    }

    [Theory]
    [InlineData(46)] // Second component XRsiz differs from the first component.
    [InlineData(47)] // Second component YRsiz differs from the first component.
    public void ImageValidationRejectsComponentTransformWithMismatchedSampling(int samplingOffset) {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] rawCodestream = payload.Skip(codestream).ToArray();
        int codingStyle = FindMarker(rawCodestream, 0xFF, 0x52);
        rawCodestream[codingStyle + 8] = 1;
        Assert.True(OfficeImageReader.TryValidateContent(rawCodestream, "scan.j2c", out _));
        Assert.Equal(1, rawCodestream[samplingOffset]);
        rawCodestream[samplingOffset] = 2;

        Assert.True(OfficeImageReader.TryIdentifyByContent(rawCodestream, "scan.j2c", out _));
        Assert.False(OfficeJpeg2000Header.TryValidateOpaquePayload(rawCodestream, out _, out _, out _));
        Assert.False(OfficeImageReader.TryValidateContent(rawCodestream, "scan.j2c", out _));
    }

    [Theory]
    [InlineData(0, 0x43)] // Quantization style 3 is reserved.
    public void ImageValidationRejectsJpeg2000CodestreamWithInvalidQuantizationDefaultField(
        int fieldOffset,
        int invalidValue) {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] rawCodestream = payload.Skip(codestream).ToArray();
        int quantization = FindMarker(rawCodestream, 0xFF, 0x5C);
        rawCodestream[quantization + 4 + fieldOffset] = (byte)invalidValue;

        Assert.False(OfficeImageReader.TryValidateContent(rawCodestream, "scan.j2c", out _));
    }

    [Fact]
    public void ImageValidationRejectsJpeg2000CodestreamWhenQuantizationDoesNotCoverCodLevels() {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] rawCodestream = payload.Skip(codestream).ToArray();
        int codingStyle = FindMarker(rawCodestream, 0xFF, 0x52);
        rawCodestream[codingStyle + 9] = 1;

        Assert.True(OfficeImageReader.TryIdentifyByContent(rawCodestream, "scan.j2c", out _));
        Assert.False(OfficeImageReader.TryValidateContent(rawCodestream, "scan.j2c", out _));
    }

    [Fact]
    public void ImageValidationRejectsJpeg2000TileCodOverrideWhenInheritedQuantizationIsTooShort() {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] rawCodestream = payload.Skip(codestream).ToArray();
        int tilePart = FindMarker(rawCodestream, 0xFF, 0x90);
        int startOfData = FindMarker(rawCodestream, 0xFF, 0x93);
        uint tilePartLength = (uint)ReadUInt32BigEndian(rawCodestream, tilePart + 6);
        byte[] tileCodingStyle = {
            0xFF, 0x52, 0x00, 0x0C,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x04, 0x04, 0x00, 0x01
        };
        byte[] malformed = rawCodestream.Take(startOfData)
            .Concat(tileCodingStyle)
            .Concat(rawCodestream.Skip(startOfData))
            .ToArray();
        WriteJpxUInt32(malformed, tilePart + 6, tilePartLength + (uint)tileCodingStyle.Length);

        Assert.True(OfficeImageReader.TryIdentifyByContent(malformed, "scan.j2c", out _));
        Assert.False(OfficeImageReader.TryValidateContent(malformed, "scan.j2c", out _));
    }

    [Fact]
    public void ImageValidationRejectsJp2WithFileTypeBoxAfterHeader() {
        byte[] payload = ReadScanJpx("rgb");
        int fileTypeLength = ReadUInt32BigEndian(payload, 12);
        byte[] reordered = payload.Take(12)
            .Concat(payload.Skip(12 + fileTypeLength))
            .Concat(payload.Skip(12).Take(fileTypeLength))
            .ToArray();

        Assert.False(OfficeImageReader.TryValidateContent(reordered, "scan.jp2", out _));
    }

    [Theory]
    [InlineData(4, 0xFF)] // Isot selects a tile outside the one-tile SIZ grid.
    [InlineData(10, 1)] // TPsot must start at zero.
    [InlineData(11, 2)] // TNsot declares a second tile-part that is absent.
    public void ImageValidationRejectsJpeg2000CodestreamWithInvalidTilePartIdentity(
        int tilePartFieldOffset,
        int invalidValue) {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] rawCodestream = payload.Skip(codestream).ToArray();
        int tilePart = FindMarker(rawCodestream, 0xFF, 0x90);
        rawCodestream[tilePart + tilePartFieldOffset] = (byte)invalidValue;

        Assert.True(OfficeImageReader.TryIdentifyByContent(rawCodestream, "scan.j2c", out _));
        Assert.False(OfficeImageReader.TryValidateContent(rawCodestream, "scan.j2c", out _));
    }

    [Theory]
    [InlineData(0x90)] // A nested SOT marker cannot begin inside bounded packet data.
    [InlineData(0x93)] // A nested SOD marker cannot begin inside bounded packet data.
    [InlineData(0xD9)] // EOC before the declared tile-part end terminates the codestream prematurely.
    public void ImageValidationRejectsPrematureMarkersInsideBoundedPacketData(int markerCode) {
        byte[] payload = ReadScanJpx("rgb");
        int codestream = FindMarker(payload, 0xFF, 0x4F, 0xFF, 0x51);
        byte[] rawCodestream = payload.Skip(codestream).ToArray();
        Assert.True(OfficeImageReader.TryValidateContent(rawCodestream, "scan.j2c", out _));
        int tilePart = FindMarker(rawCodestream, 0xFF, 0x90);
        int startOfData = FindMarker(rawCodestream, 0xFF, 0x93);
        int tilePartEnd = tilePart + ReadUInt32BigEndian(rawCodestream, tilePart + 6);
        Assert.True(startOfData + 4 <= tilePartEnd);
        rawCodestream[startOfData + 2] = 0xFF;
        rawCodestream[startOfData + 3] = (byte)markerCode;

        Assert.False(OfficeJpeg2000Header.TryValidateOpaquePayload(rawCodestream, out _, out _, out _));
        Assert.False(OfficeImageReader.TryValidateContent(rawCodestream, "scan.j2c", out _));
    }

    [Fact]
    public void Jpeg2000ValidationHonorsCancellation() {
        byte[] payload = ReadScanJpx("rgb");
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => {
            OfficeImageReader.TryValidateContent(payload, "scan.jp2", cancellation.Token, out _);
        });
    }

    private static int FindMarker(byte[] bytes, params byte[] marker) =>
        Enumerable.Range(0, bytes.Length - marker.Length + 1).First(index =>
            marker.Select((value, markerIndex) => bytes[index + markerIndex] == value).All(static match => match));

    private static int ReadUInt32BigEndian(byte[] bytes, int offset) =>
        (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    private static byte[] CreateJp2Box(string type, byte[] contents) {
        Assert.Equal(4, type.Length);
        byte[] box = new byte[8 + contents.Length];
        WriteUInt32BigEndian(box, 0, box.Length);
        System.Text.Encoding.ASCII.GetBytes(type, 0, 4, box, 4);
        Buffer.BlockCopy(contents, 0, box, 8, contents.Length);
        return box;
    }

    private static byte[] InsertJp2ResolutionBox(byte[] payload, byte[] resolutionBox) {
        int headerType = FindMarker(payload, (byte)'j', (byte)'p', (byte)'2', (byte)'h');
        int headerStart = headerType - 4;
        int headerEnd = headerStart + ReadUInt32BigEndian(payload, headerStart);
        byte[] updated = payload.Take(headerEnd).Concat(resolutionBox).Concat(payload.Skip(headerEnd)).ToArray();
        WriteUInt32BigEndian(updated, headerStart, headerEnd - headerStart + resolutionBox.Length);
        return updated;
    }

    private static void WriteUInt32BigEndian(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }

    private static byte[] ReadScanJpx(string mode) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,
        "Pdf", "Fixtures", "Interoperability", "Scans", "red-" + mode + ".jp2"));

    private sealed class ScanJpxCodec : IOfficeRasterImageCodec {
        private readonly byte[] _expected;
        private readonly string _expectedContentType;
        internal ScanJpxCodec(byte[] expected) {
            _expected = expected;
            _expectedContentType = OfficeJpeg2000Header.IsJp2Container(expected) ? "image/jp2" : "image/j2c";
        }
        internal int Calls { get; private set; }
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            Assert.Equal(_expectedContentType, contentType);
            Assert.Equal(_expected, encodedBytes);
            Calls++;
            Assert.Equal(1, Calls); // A one-shot codec must suffice for a single image placement.
            image = new OfficeRasterImage(1, 1, OfficeColor.Red);
            return true;
        }
    }
}

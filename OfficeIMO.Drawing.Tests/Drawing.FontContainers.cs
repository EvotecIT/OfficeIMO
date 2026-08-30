using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingFontContainerTests {
    [Fact]
    public void OfficeFontContainerDecoder_RoundTripsCompressedWoffIntoReusableOpenType() {
        byte[] source = ManagedTextShapingTestAssets.CreateFont('A', 0x1F600);
        int headOffset = FindTableOffset(source, "head");
        WriteUInt32(source, headOffset + 8, 0xBADB455D);
        byte[] woff = ManagedTextShapingTestAssets.CreateWoff(source);

        bool decoded = OfficeFontContainerDecoder.TryDecodeToOpenType(
            woff,
            out byte[] openType,
            out OfficeFontContainerFormat format,
            out string? error);

        Assert.True(decoded, error);
        Assert.Equal(OfficeFontContainerFormat.Woff, format);
        Assert.NotSame(source, openType);
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(openType));
        Assert.True(font.HasGlyphs("A" + char.ConvertFromUtf32(0x1F600)));
        Assert.Equal(0xB1B0AFBA, CalculateChecksum(openType));
        Assert.NotEqual(0U, ReadUInt32(openType, FindTableOffset(openType, "head") + 8));
        var fonts = new OfficeFontFaceCollection().Add("WOFF Demo", woff);
        Assert.Single(fonts.Faces);
        Assert.Equal(OfficeFontContainerFormat.OpenType, OfficeFontContainerDecoder.Detect(fonts.Faces[0].Data));
    }

    [Fact]
    public void OfficeFontContainerDecoder_RejectsMalformedOrOversizedContainersWithoutPartialOutput() {
        byte[] source = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] woff = ManagedTextShapingTestAssets.CreateWoff(source);
        int firstTableOffset = (woff[48] << 24) | (woff[49] << 16) | (woff[50] << 8) | woff[51];
        woff[firstTableOffset] ^= 0x7F;

        Assert.False(OfficeFontContainerDecoder.TryDecodeToOpenType(
            woff,
            out byte[] decoded,
            out OfficeFontContainerFormat format,
            out string? error));
        Assert.Equal(OfficeFontContainerFormat.Woff, format);
        Assert.Empty(decoded);
        Assert.False(string.IsNullOrWhiteSpace(error));

        byte[] valid = ManagedTextShapingTestAssets.CreateWoff(source);
        Assert.False(OfficeFontContainerDecoder.TryDecodeToOpenType(
            valid,
            source.Length - 1,
            out decoded,
            out format,
            out error));
        Assert.Empty(decoded);
        Assert.Contains("limit", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OfficeFontContainerDecoder_DetectsWoff2AndReportsItsCurrentBoundary() {
        byte[] woff2 = { 0x77, 0x4F, 0x46, 0x32, 0, 0, 0, 0 };

        Assert.Equal(OfficeFontContainerFormat.Woff2, OfficeFontContainerDecoder.Detect(woff2));
        Assert.False(OfficeFontContainerDecoder.TryDecodeToOpenType(
            woff2,
            out byte[] decoded,
            out OfficeFontContainerFormat format,
            out string? error));
        Assert.Equal(OfficeFontContainerFormat.Woff2, format);
        Assert.Empty(decoded);
        Assert.Contains("WOFF 2", error, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeFontFaceCollection_UsesOptionalProviderForBoundedWoff2Programs() {
        byte[] woff2 = { 0x77, 0x4F, 0x46, 0x32, 1, 2, 3, 4 };
        var provider = new TestFontProgramProvider(decodedByteCount: 32);
        var fonts = new OfficeFontFaceCollection {
            FontProgramProvider = provider
        };

        Assert.True(fonts.TryAdd("Provider Demo", woff2));

        OfficeFontFace face = Assert.Single(fonts.Faces);
        Assert.Equal(OfficeFontContainerFormat.Woff2, face.ContainerFormat);
        Assert.False(face.CanEmbedAsStaticPdfFont);
        Assert.Equal("A", Assert.Single(fonts.PlanFallbackRuns("A", "Provider Demo")).Text);
        Assert.True(fonts.TryMeasureText(
            "A",
            12D,
            "Provider Demo",
            OfficeFontStyle.Regular,
            out double width));
        Assert.Equal(42D, width);
        Assert.Same(provider, fonts.Clone().FontProgramProvider);
        Assert.Equal(OfficeFontContainerFormat.Woff2, provider.LastRequest!.ContainerFormat);
        Assert.Equal("Provider Demo", provider.LastRequest.FamilyName);
    }

    [Fact]
    public void OfficeFontFaceCollection_RejectsProviderProgramsThatExceedTheDecodedLimit() {
        byte[] woff2 = { 0x77, 0x4F, 0x46, 0x32, 1, 2, 3, 4 };
        var fonts = new OfficeFontFaceCollection {
            FontProgramProvider = new TestFontProgramProvider(decodedByteCount: 65)
        };

        Assert.False(fonts.TryAddBounded(
            "Provider Demo",
            woff2,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            maximumDecodedBytes: 64,
            out int decodedBytes,
            out string? error));

        Assert.Equal(0, decodedBytes);
        Assert.Contains("limit", error, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(fonts.Faces);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OfficeFontFaceCollection_CountsProviderAndFaceBuffersIndependently(bool includeStaticSnapshot) {
        byte[] woff2 = { 0x77, 0x4F, 0x46, 0x32, 1, 2, 3, 4 };
        var fonts = new OfficeFontFaceCollection {
            FontProgramProvider = new TestFontProgramProvider(
                decodedByteCount: 57,
                staticOpenTypeData: includeStaticSnapshot ? woff2 : null)
        };

        Assert.False(fonts.TryAddBounded(
            "Provider Demo",
            woff2,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            maximumDecodedBytes: 64,
            out int rejectedBytes,
            out string? rejectedError));
        Assert.Equal(0, rejectedBytes);
        Assert.Contains("limit", rejectedError, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(fonts.Faces);

        Assert.True(fonts.TryAddBounded(
            "Provider Demo",
            woff2,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            maximumDecodedBytes: 65,
            out int acceptedBytes,
            out string? acceptedError), acceptedError);
        Assert.Equal(65, acceptedBytes);
        Assert.Single(fonts.Faces);
    }

    [Fact]
    public void OfficeFontFaceCollection_RejectsProviderAndFaceBufferTotalsBeyondInt32() {
        byte[] woff2 = { 0x77, 0x4F, 0x46, 0x32, 1, 2, 3, 4 };
        var fonts = new OfficeFontFaceCollection {
            FontProgramProvider = new TestFontProgramProvider(int.MaxValue)
        };

        Assert.False(fonts.TryAddBounded(
            "Provider Demo",
            woff2,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            maximumDecodedBytes: int.MaxValue,
            out int decodedBytes,
            out string? error));

        Assert.Equal(0, decodedBytes);
        Assert.Contains("limit", error, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(fonts.Faces);
    }

    [Fact]
    public void TrueTypeShapingDataReturnsIndependentSnapshots() {
        byte[] source = ManagedTextShapingTestAssets.CreateFont('A');
        OfficeFontFace face = Assert.Single(new OfficeFontFaceCollection().Add("Snapshot Demo", source).Faces);

        byte[] first = face.Program.GetFontDataForShaping();
        byte original = first[0];
        first[0] ^= 0xFF;
        byte[] second = face.Program.GetFontDataForShaping();

        Assert.NotSame(first, second);
        Assert.Equal(original, second[0]);
    }

    [Fact]
    public void OfficeFontFaceCollection_CountsBuiltInTrueTypeAndFaceBuffersIndependently() {
        byte[] source = ManagedTextShapingTestAssets.CreateFont('A');
        var fonts = new OfficeFontFaceCollection();

        Assert.False(fonts.TryAddBounded(
            "Bounded TrueType",
            source,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            maximumDecodedBytes: source.Length * 2 - 1,
            out int rejectedBytes,
            out string? rejectedError));
        Assert.Equal(0, rejectedBytes);
        Assert.Contains("limit", rejectedError, StringComparison.OrdinalIgnoreCase);

        Assert.True(fonts.TryAddBounded(
            "Bounded TrueType",
            source,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            maximumDecodedBytes: source.Length * 2,
            out int acceptedBytes,
            out string? acceptedError), acceptedError);
        Assert.Equal(source.Length * 2, acceptedBytes);
        Assert.Single(fonts.Faces);
    }

    private sealed class TestFontProgramProvider : IOfficeFontProgramProvider {
        private readonly int _decodedByteCount;
        private readonly byte[]? _staticOpenTypeData;

        internal TestFontProgramProvider(int decodedByteCount, byte[]? staticOpenTypeData = null) {
            _decodedByteCount = decodedByteCount;
            _staticOpenTypeData = staticOpenTypeData;
        }

        internal OfficeFontProgramLoadRequest? LastRequest { get; private set; }

        public OfficeFontProgramLoadResult? TryLoad(OfficeFontProgramLoadRequest request) {
            LastRequest = request;
            return new OfficeFontProgramLoadResult(new TestFontProgram(), _decodedByteCount, _staticOpenTypeData);
        }
    }

    private sealed class TestFontProgram : IOfficeFontProgram {
        public string Fingerprint => "test-font-program-v1";
        public string? DisplayName => "Test font program";
        public int? CollectionIndex => null;
        public int UnitsPerEm => 1000;
        public bool IsOpenTypeCff => false;
        public bool ProvidesComplexTextLayout => true;
        public double LineSpacingRatio => 1D;
        public byte[] GetFontDataForShaping() => new byte[] { 1 };
        public bool HasGlyphs(string text) => !string.IsNullOrEmpty(text);
        public double Measure(string text, double fontSize) => text.Length * 42D;
        public IReadOnlyList<double> MeasureTextElements(IReadOnlyList<string> elements, double fontSize) =>
            elements.Select(element => element.Length * 42D).ToArray();
        public double LineHeight(double fontSize) => fontSize;
        public List<List<OfficePoint>> GetTextContours(string text, double x, double y, double fontSize) =>
            new List<List<OfficePoint>>();
        public bool TryGetGlyphMetrics(int scalar, out int glyphId, out int advanceWidth) {
            glyphId = scalar;
            advanceWidth = 42;
            return true;
        }
        public double MeasureShapedText(string text, OfficeTextShapingResult result, double fontSize) =>
            Measure(text, fontSize);
        public List<List<OfficePoint>> GetShapedTextContours(
            string text,
            OfficeTextShapingResult result,
            double x,
            double y,
            double fontSize) => GetTextContours(text, x, y, fontSize);
    }

    private static int FindTableOffset(byte[] font, string tag) {
        int tableCount = (font[4] << 8) | font[5];
        for (int index = 0; index < tableCount; index++) {
            int record = 12 + index * 16;
            if (font[record] == tag[0] && font[record + 1] == tag[1]
                && font[record + 2] == tag[2] && font[record + 3] == tag[3]) {
                return checked((int)ReadUInt32(font, record + 8));
            }
        }
        throw new InvalidOperationException("The test font has no " + tag + " table.");
    }

    private static uint CalculateChecksum(byte[] data) {
        uint checksum = 0;
        for (int offset = 0; offset < data.Length; offset += 4) {
            uint value = (uint)data[offset] << 24;
            if (offset + 1 < data.Length) value |= (uint)data[offset + 1] << 16;
            if (offset + 2 < data.Length) value |= (uint)data[offset + 2] << 8;
            if (offset + 3 < data.Length) value |= data[offset + 3];
            checksum = unchecked(checksum + value);
        }
        return checksum;
    }

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24)
        | ((uint)data[offset + 1] << 16)
        | ((uint)data[offset + 2] << 8)
        | data[offset + 3];

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}

using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Official AOM-coded streams provide an independent oracle for symbol values, adaptation and tile termination.</summary>
public sealed class DrawingAv1SymbolTests {
    public static IEnumerable<object[]> ReferenceCases() {
        using var json = JsonDocument.Parse(File.ReadAllText(FixturePath()));
        foreach (var c in json.RootElement.GetProperty("cases").EnumerateArray())
            yield return new object[] { c.GetProperty("symbols").GetInt32(), c.GetProperty("updates").GetBoolean(),
                c.GetProperty("operations").GetInt32(), c.GetProperty("hex").GetString()!,
                c.GetProperty("initialCdf").EnumerateArray().Select(v => v.GetInt32()).ToArray(),
                c.GetProperty("finalCdf").EnumerateArray().Select(v => v.GetInt32()).ToArray() };
    }

    [Theory]
    [MemberData(nameof(ReferenceCases))]
    public void OfficialAomStreamsDecodeSymbolsLiteralsAndAdaptiveCdfs(int symbols, bool updates,
        int operations, string hex, int[] initialCdf, int[] finalCdf) {
        byte[] encoded = Hex(hex);
        // Nonzero neighboring bytes must not contribute lookahead or termination bits to this tile.
        byte[] surrounding = Enumerable.Repeat((byte)255, encoded.Length + 6).ToArray();
        Buffer.BlockCopy(encoded, 0, surrounding, 3, encoded.Length);
        long budget = operations + (operations + 4) / 5 + (operations + 31) / 32 * 31;
        var reader = new OfficeAv1SymbolReader(surrounding, 3, encoded.Length, budget, !updates);
        int[] cdf = (int[])initialCdf.Clone();
        for (int i = 0; i < operations; i++) {
            Assert.Equal((i * 13 + i / 7) % symbols, reader.ReadSymbol(cdf));
            if (i % 5 == 0) Assert.Equal(((i / 5) & 1) != 0, reader.ReadBool());
            if (i % 32 == 0) Assert.Equal((i * 139 + 71) & int.MaxValue, reader.ReadLiteral(31));
        }
        reader.Finish();
        Assert.Equal(finalCdf, cdf);
    }

    [Fact]
    public void ShortNativeBooleanStreamsHonorTheirOwnTrailingBits() {
        using var json = JsonDocument.Parse(File.ReadAllText(FixturePath()));
        foreach (var c in json.RootElement.GetProperty("booleanCases").EnumerateArray()) {
            string bits = c.GetProperty("bits").GetString()!;
            byte[] payload = Hex(c.GetProperty("hex").GetString()!);
            var reader = new OfficeAv1SymbolReader(payload, 0, payload.Length, bits.Length, true);
            foreach (char bit in bits) Assert.Equal(bit == '1', reader.ReadBool());
            reader.Finish();
        }
    }

    [Fact]
    public void FrozenAvifTilePrefixesMatchIndependentNativePartitionSymbols() {
        using var json = JsonDocument.Parse(File.ReadAllText(FixturePath()));
        foreach (var c in json.RootElement.GetProperty("framePrefixes").EnumerateArray()) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif",
                c.GetProperty("name").GetString()! + ".avif"));
            var options = new OfficeRasterDecodeOptions();
            Assert.True(OfficeAvifContainerReader.TryRead(bytes, options, out var container));
            var item = c.GetProperty("alpha").GetBoolean() ? container!.Alpha! : container!.Color;
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes, item, options, out var sequence));
            Assert.True(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence!, options, out var frame));
            var tile = frame!.Tiles[0];
            Assert.Equal(c.GetProperty("offset").GetInt32(), tile.Offset);
            Assert.Equal(c.GetProperty("length").GetInt32(), tile.Length);
            // AV1 Default_Partition_W64_Cdf[0]; native AOM independently decoded this tile slice.
            int[] cdf = { 20137, 21547, 23078, 29566, 29837, 30261, 30524, 30892, 31724, 32768, 0 };
            var reader = new OfficeAv1SymbolReader(bytes, tile.Offset, tile.Length, 1, frame.DisableCdfUpdate);
            Assert.Equal(c.GetProperty("symbol").GetInt32(), reader.ReadSymbol(cdf));
            Assert.Equal(c.GetProperty("finalCdf").EnumerateArray().Select(v => v.GetInt32()).ToArray(), cdf);
            // This is prefix proof only: Finish requires the complete tile grammar, which is not decoded here.
        }
    }

    [Fact]
    public void TruncatedTileCannotBorrowFollowingBytesOrUnlimitedSyntheticZeros() {
        var reader = new OfficeAv1SymbolReader(new byte[] { 0, 128, 255 }, 0, 1, 1000, true);
        Assert.Throws<FormatException>(() => {
            // AV1 permits at most 14 synthetic lookahead bits, not an arbitrary zero tail.
            for (int i = 0; i < 32; i++) reader.ReadBool();
            reader.Finish();
        });
        Assert.Throws<FormatException>(() => reader.ReadBool());
        Assert.Throws<FormatException>(() => new OfficeAv1SymbolReader(new byte[] { 32 }, 0, 2, 1, true));
    }

    [Theory]
    [InlineData(0)] // Removed mandatory trailing one from native "0" stream 0x20.
    [InlineData(33)] // Added a nonzero trailing padding bit.
    public void InvalidTileTerminationCannotBeAcceptedOrResumed(byte malformed) {
        var reader = new OfficeAv1SymbolReader(new[] { malformed }, 0, 1, 10, true);
        Assert.False(reader.ReadBool());
        Assert.Throws<FormatException>(() => reader.Finish());
        Assert.Throws<FormatException>(() => reader.ReadBool());
    }

    [Fact]
    public void SymbolBudgetAndCancellationStopBeforeFurtherAdaptation() {
        var reader = new OfficeAv1SymbolReader(Hex("2b80"), 0, 2, 2, false);
        Assert.False(reader.ReadBool());
        Assert.False(reader.ReadBool());
        Assert.Throws<FormatException>(() => reader.ReadBool());
        using var cancellation = new CancellationTokenSource();
        reader = new OfficeAv1SymbolReader(Hex("2b80"), 0, 2, 10, false, cancellation.Token);
        int[] cdf = { 16384, 32768, 0 };
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => reader.ReadSymbol(cdf));
        Assert.Equal(new[] { 16384, 32768, 0 }, cdf);
        Assert.Throws<OperationCanceledException>(() => new OfficeAv1SymbolReader(Hex("20"), 0, 1, 10, true, cancellation.Token));
    }

    private static string FixturePath() => Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif", "entropy-reference.json");
    private static byte[] Hex(string hex) {
        var result = new byte[hex.Length / 2];
        for (int i = 0; i < result.Length; i++) result[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
        return result;
    }
}

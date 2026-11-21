using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// JBIG2 decoding (ITU-T T.88) of independently encoded streams against the encoder's source image or jbig2dec.
/// See OfficeIMO.TestAssets/Images/Jbig2/PROVENANCE.md.
/// </summary>
public sealed class Jbig2DecoderTests {
    private static readonly string Folder = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Images", "Jbig2");

    public static TheoryData<string> Fixtures() {
        var data = new TheoryData<string>();
        foreach (JsonElement entry in ReadManifest().GetProperty("files").EnumerateArray()) data.Add(entry.GetProperty("file").GetString()!);
        return data;
    }

    [Theory]
    [MemberData(nameof(Fixtures))]
    public void DecodesFixtureLikeTheReference(string file) {
        JsonElement entry = ReadManifest().GetProperty("files").EnumerateArray().Single(item => item.GetProperty("file").GetString() == file);
        byte[] bytes = File.ReadAllBytes(Path.Combine(Folder, file));
        Assert.Equal(entry.GetProperty("sha256").GetString(), Sha256(bytes));
        int width = entry.GetProperty("width").GetInt32(), height = entry.GetProperty("height").GetInt32();

        byte[] rows = Jbig2Decoder.Decode(ToEmbedded(bytes), null, width, height, 64 * 1024 * 1024, CancellationToken.None);
        for (int index = 0; index < rows.Length; index++) rows[index] = (byte)~rows[index];
        ClearPadding(rows, width, height);
        Assert.Equal(entry.GetProperty("rowsSha256").GetString(), Sha256(rows));
        Assert.Null(Jbig2Decoder.FindUnsupportedFeature(ToEmbedded(bytes)));
    }

    // Untrusted input: truncated and corrupted streams must fail with InvalidDataException or decode within budget.
    [Theory]
    [InlineData("bitmap-symbol-symhuffrefine-textrefine.jbig2")]
    [InlineData("bitmap-halftone-10bpp-mmr.jbig2")]
    [InlineData("bitmap-symbol-texthuffrefinecustom.jbig2")]
    [InlineData("bitmap-refine-template1-tpgron.jbig2")]
    public void MalformedInputFailsClosed(string file) {
        byte[] original = ToEmbedded(File.ReadAllBytes(Path.Combine(Folder, file)));
        var random = new Random(file.Length);
        for (int round = 0; round < 400; round++) {
            byte[] mutated = round < 100 ? original.Take(original.Length * round / 100).ToArray() : (byte[])original.Clone();
            if (round >= 100) {
                for (int change = 0; change <= random.Next(6); change++) mutated[random.Next(mutated.Length)] = (byte)random.Next(256);
            }
            try {
                Jbig2Decoder.Decode(mutated, null, 399, 400, 4 * 1024 * 1024, CancellationToken.None);
            } catch (InvalidDataException) {
            } catch (NotSupportedException) {
            }
        }
    }

    [Fact]
    public void DecodingObservesCancellation() {
        byte[] embedded = ToEmbedded(File.ReadAllBytes(Path.Combine(Folder, "bitmap-halftone.jbig2")));
        Assert.ThrowsAny<OperationCanceledException>(() =>
            Jbig2Decoder.Decode(embedded, null, 399, 400, 64 * 1024 * 1024, new CancellationToken(true)));
    }

    // Annex D file organisations (sequential or random access) to the embedded PDF organisation for page 1 (D.3).
    private static byte[] ToEmbedded(byte[] file) {
        int position = 8;
        int flags = file[position++];
        if ((flags & 2) == 0) position += 4;
        bool sequential = (flags & 1) != 0;
        var output = new MemoryStream();
        var headers = new List<(int Start, int Length, uint DataLength, uint Page)>();
        while (position < file.Length) {
            int start = position;
            uint number = ReadUInt32(file, position);
            position += 4;
            int segmentFlags = file[position++];
            int referred = file[position] >> 5;
            if (referred == 7) {
                referred = (int)(ReadUInt32(file, position) & 0x1FFFFFFF);
                position += 4 + (referred + 8) / 8;
            } else {
                position++;
            }
            position += referred * (number <= 256 ? 1 : number <= 65536 ? 2 : 4);
            uint page = (segmentFlags & 0x40) != 0 ? ReadUInt32(file, position) : file[position];
            position += (segmentFlags & 0x40) != 0 ? 4 : 1;
            uint dataLength = ReadUInt32(file, position);
            position += 4;
            var header = (start, position - start, dataLength, page);
            if (sequential) {
                if (page <= 1) {
                    output.Write(file, start, position - start);
                    output.Write(file, position, (int)dataLength);
                }
                position += (int)dataLength;
            } else {
                headers.Add(header);
            }
            if ((segmentFlags & 0x3F) == 51) break;
        }
        foreach (var header in headers) {
            if (header.Page <= 1) {
                output.Write(file, header.Start, header.Length);
                output.Write(file, position, (int)header.DataLength);
            }
            position += (int)header.DataLength;
        }
        return output.ToArray();
    }

    private static void ClearPadding(byte[] rows, int width, int height) {
        int stride = (width + 7) / 8, bits = width & 7;
        if (bits == 0) return;
        for (int y = 0; y < height; y++) rows[y * stride + stride - 1] &= (byte)(0xFF << (8 - bits));
    }

    private static uint ReadUInt32(byte[] data, int offset) =>
        (uint)(data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3]);

    private static string Sha256(byte[] data) {
        using var hash = SHA256.Create();
        return BitConverter.ToString(hash.ComputeHash(data)).Replace("-", string.Empty).ToLowerInvariant();
    }

    private static JsonElement ReadManifest() =>
        JsonDocument.Parse(File.ReadAllText(Path.Combine(Folder, "manifest.json"))).RootElement;
}

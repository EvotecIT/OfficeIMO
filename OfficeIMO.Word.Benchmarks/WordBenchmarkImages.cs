using System.Buffers.Binary;
using System.Text;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Distinct, valid PNG payloads derived from one source image so no library can deduplicate them.</summary>
internal static class WordBenchmarkImages {
    internal const int WidthPixels = 287;
    internal const int HeightPixels = 67;
    private static byte[]? _source;
    private static readonly Dictionary<int, byte[]> Variants = new();

    private static byte[] Source => _source ??= File.ReadAllBytes(
        Path.Combine(AppContext.BaseDirectory, "benchmark-logo.png"));

    internal static byte[] Image(int index) {
        lock (Variants) {
            if (!Variants.TryGetValue(index, out byte[]? image)) {
                image = WithTextChunk(Source, "Comment", "Benchmark image " + index.ToString(System.Globalization.CultureInfo.InvariantCulture));
                Variants[index] = image;
            }
            return image;
        }
    }

    /// <summary>Inserts a tEXt chunk after IHDR. The pixels stay identical while the file bytes differ.</summary>
    private static byte[] WithTextChunk(byte[] png, string keyword, string text) {
        const int signatureAndHeader = 8 + 4 + 4 + 13 + 4;
        if (png.Length < signatureAndHeader || Encoding.ASCII.GetString(png, 12, 4) != "IHDR") {
            throw new InvalidDataException("The benchmark image is not a PNG with a leading IHDR chunk.");
        }
        byte[] data = Encoding.Latin1.GetBytes(keyword + "\0" + text);
        byte[] type = Encoding.ASCII.GetBytes("tEXt");
        var output = new byte[png.Length + 12 + data.Length];
        Buffer.BlockCopy(png, 0, output, 0, signatureAndHeader);
        int offset = signatureAndHeader;
        BinaryPrimitives.WriteInt32BigEndian(output.AsSpan(offset), data.Length);
        Buffer.BlockCopy(type, 0, output, offset + 4, 4);
        Buffer.BlockCopy(data, 0, output, offset + 8, data.Length);
        BinaryPrimitives.WriteUInt32BigEndian(output.AsSpan(offset + 8 + data.Length), Crc32(output.AsSpan(offset + 4, 4 + data.Length)));
        Buffer.BlockCopy(png, signatureAndHeader, output, offset + 12 + data.Length, png.Length - signatureAndHeader);
        return output;
    }

    private static uint Crc32(ReadOnlySpan<byte> bytes) {
        uint crc = 0xFFFFFFFFu;
        foreach (byte value in bytes) {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320u : crc >> 1;
        }
        return ~crc;
    }
}

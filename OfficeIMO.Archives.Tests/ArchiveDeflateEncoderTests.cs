using System.IO.Compression;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveDeflateEncoderTests {
    public static TheoryData<string> Shapes => new() { "empty", "one", "high", "text", "random", "runs", "mixed" };

    [Theory]
    [MemberData(nameof(Shapes))]
    public void EncodedStreamsDecodeExactlyWithTheRuntimeAndTheStrictValidator(string shape) {
        byte[] input = Sample(shape);
        byte[] encoded = Encode(input, input.Length == 0 ? 1 : input.Length);

        Assert.True(OfficeDeflateStreamValidator.TryValidateExact(encoded, 0, encoded.Length, input.Length));
        using var inflater = new DeflateStream(new MemoryStream(encoded), CompressionMode.Decompress);
        using var output = new MemoryStream();
        inflater.CopyTo(output);
        Assert.Equal(input, output.ToArray());
        if (shape == "runs") Assert.True(encoded.Length < input.Length / 50);
        if (shape == "random") Assert.True(encoded.Length <= input.Length + input.Length / 1000 + 16);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(7)]
    [InlineData(65536)]
    public void OutputDoesNotDependOnWriteChunking(int chunk) {
        byte[] input = Sample("mixed");
        Assert.Equal(Encode(input, input.Length), Encode(input, chunk));
    }

    internal static byte[] Encode(byte[] input, int chunk) {
        using var output = new MemoryStream();
        var encoder = new ArchiveDeflateEncoder(output);
        for (int offset = 0; offset < input.Length; offset += chunk) encoder.Write(input, offset, Math.Min(chunk, input.Length - offset));
        encoder.Finish();
        return output.ToArray();
    }

    internal static byte[] Sample(string shape) {
        var random = new Random(1234);
        switch (shape) {
            case "empty": return Array.Empty<byte>();
            case "one": return new byte[] { 42 };
            // Short input takes the fixed-code block, where bytes 144-255 use the nine-bit codes.
            case "high": return Encoding.UTF8.GetBytes("Zażółć gęślą jaźń.\n");
            case "text": return Encoding.UTF8.GetBytes(string.Concat(Enumerable.Range(0, 2000).Select(index => "Line " + index + ": zażółć gęślą jaźń.\n")));
            case "random": { var bytes = new byte[300_000]; random.NextBytes(bytes); return bytes; }
            case "runs": return Enumerable.Range(0, 400_000).Select(index => (byte)(index / 5000)).ToArray();
            default: {
                // Text, incompressible spans and long repeats across several 32 KiB windows and block boundaries.
                var output = new MemoryStream();
                for (int round = 0; round < 40; round++) {
                    byte[] text = Sample("text");
                    output.Write(text, 0, random.Next(1, text.Length));
                    var noise = new byte[random.Next(1, 40_000)]; random.NextBytes(noise); output.Write(noise, 0, noise.Length);
                    var zeros = new byte[random.Next(1, 70_000)]; output.Write(zeros, 0, zeros.Length);
                    for (int repeat = random.Next(300, 3000); repeat > 0; repeat--) output.WriteByte((byte)(repeat % 3));
                }
                return output.ToArray();
            }
        }
    }
}

using System.Buffers.Binary;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveCompoundStreamLengthContractTests {
    [Theory]
    [InlineData(1)]
    [InlineData(63)]
    [InlineData(64)]
    [InlineData(65)]
    [InlineData(4_095)]
    [InlineData(4_096)]
    [InlineData(4_097)]
    [InlineData(10_000)]
    public void FullReaderReturnsEachStreamAtExactlyItsDeclaredLength(int length) {
        // Neighbouring mini and regular streams make the probe's last sector share space with other data.
        byte[] probe = CreatePayload(length, seed: 7);
        byte[] file = OfficeCompoundFileWriter.Write(new[] {
            new OfficeCompoundStream("MiniNeighbour", CreatePayload(700, seed: 3)),
            new OfficeCompoundStream("SizeProbe", probe),
            new OfficeCompoundStream("RegularNeighbour", CreatePayload(5_000, seed: 5))
        });

        Assert.True(OfficeCompoundFileReader.TryRead(file, out Dictionary<string, byte[]> streams, out string? error), error);
        Assert.Equal(probe, streams["SizeProbe"]);
        Assert.Equal(CreatePayload(700, seed: 3), streams["MiniNeighbour"]);
        Assert.Equal(CreatePayload(5_000, seed: 5), streams["RegularNeighbour"]);
    }

    [Theory]
    [InlineData(100, 300, "Mini sector chain is shorter than the declared stream size.")]
    [InlineData(5_000, 5_600, "Compound file sector chain is shorter than the declared stream size.")]
    public void FullReaderRejectsAStreamDeclaredLongerThanItsSectorChain(int length, int declaredLength, string message) {
        byte[] file = OfficeCompoundFileWriter.Write(new[] {
            new OfficeCompoundStream("SizeProbe", CreatePayload(length, seed: 11))
        });
        int entry = file.AsSpan().IndexOf(Encoding.Unicode.GetBytes("SizeProbe"));
        Assert.True(entry >= 0);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(entry + 120, 4), (uint)declaredLength);

        Assert.False(OfficeCompoundFileReader.TryRead(file, out Dictionary<string, byte[]> _, out string? error));
        Assert.Contains(message, error);
    }

    private static byte[] CreatePayload(int length, int seed) {
        var payload = new byte[length];
        for (int index = 0; index < payload.Length; index++) {
            payload[index] = unchecked((byte)((index * 31) + seed));
        }

        return payload;
    }
}

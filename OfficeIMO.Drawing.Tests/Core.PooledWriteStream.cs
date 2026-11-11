#if NET8_0_OR_GREATER
using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PooledWriteStreamTests {
    [Theory]
    [InlineData(0)]
    [InlineData(65535)]
    [InlineData(65536)]
    [InlineData(65537)]
    [InlineData(131073)]
    public void SequentialRangesAndBytesProduceDetachedOutputAcrossChunkBoundaries(int length) {
        var source = new byte[length + 4];
        new Random(64).NextBytes(source);
        int first = length / 3;
        byte[] expected = source.Skip(2).Take(first).Concat(new byte[] { 0xFA })
            .Concat(source.Skip(2 + first).Take(length - first)).ToArray();
        var output = new OfficePooledWriteStream();
        output.Write(source, 2, first);
        output.WriteByte(0xFA);
        output.Write(source, 2 + first, length - first);
        Assert.Equal(length + 1, output.Length);
        Assert.Equal(output.Length, output.Position);
        Array.Clear(source, 0, source.Length);
        byte[] detached = output.ToArray();
        Assert.Equal(expected, detached);
        output.WriteByte(0xFB);
        Assert.Equal(expected, detached);
        output.Dispose();
        output.Dispose();
        Assert.Equal(expected, detached);
        Assert.False(output.CanWrite);
        Assert.Throws<ObjectDisposedException>(() => output.WriteByte(0));
    }
}
#endif

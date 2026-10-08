using System;
using System.IO;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingPortableMapCancellationTests {
    [Fact]
    public void PrecancelledPortableMapRequestsThrowAndLeaveInputIntact() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        foreach (byte[] bytes in Fixtures()) {
            byte[] original = (byte[])bytes.Clone();
            var options = new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token };
            OperationCanceledException error = Assert.Throws<OperationCanceledException>(() =>
                OfficeRasterImageDecoder.TryDecode(bytes, options, out _, out _));
            Assert.Equal(cancellation.Token, error.CancellationToken);
            Assert.Equal(original, bytes);
            Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.Black, image!.GetPixel(0, 0));
        }
    }

    [Fact]
    public void SinkCancellationAfterPbmHeaderLeavesDestinationOpenAndSourceIntact() {
        var image = new OfficeRasterImage(17, 2, OfficeColor.Black);
        byte[] original = image.GetPixels();
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelAfterHeaderStream(cancellation);
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageEncoder.EncodeTo(
            image, OfficeImageExportFormat.Pbm, stream, null, 4096, cancellation.Token));
        Assert.Equal(Encoding.ASCII.GetBytes("P4\n17 2\n"), stream.ToArray());
        Assert.Equal(original, image.GetPixels());
        Assert.True(stream.CanWrite);
    }

    private static byte[][] Fixtures() => new[] {
        Encoding.ASCII.GetBytes("P1\n1 1\n1\n"),
        Encoding.ASCII.GetBytes("P2\n1 1\n255\n0\n"),
        Encoding.ASCII.GetBytes("P3\n1 1\n255\n0 0 0\n"),
        Raw("P4\n1 1\n", new byte[] { 0x80 }),
        Raw("P5\n1 1\n255\n", new byte[] { 0 }),
        Raw("P6\n1 1\n255\n", new byte[] { 0, 0, 0 })
    };

    private static byte[] Raw(string header, byte[] samples) {
        byte[] prefix = Encoding.ASCII.GetBytes(header);
        var result = new byte[prefix.Length + samples.Length];
        Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
        Buffer.BlockCopy(samples, 0, result, prefix.Length, samples.Length);
        return result;
    }

    private sealed class CancelAfterHeaderStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        internal CancelAfterHeaderStream(CancellationTokenSource cancellation) => _cancellation = cancellation;
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            _cancellation.Cancel();
        }
    }
}

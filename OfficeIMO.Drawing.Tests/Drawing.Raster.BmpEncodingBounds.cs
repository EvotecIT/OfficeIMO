using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingBmpEncodingBoundsTests {
    [Theory]
    [InlineData(1, 1)]
    [InlineData(4097, 2)]
    [InlineData(17, 3)]
    public void BmpRoutesPreserveStraightAlphaAndBottomUpPixels(int width, int height) {
        OfficeRasterImage image = Pattern(width, height);
        byte[] before = image.GetPixels();
        var options = new OfficeRasterEncodingOptions { DpiX = 144D, DpiY = 120D };
        byte[] encoded = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Bmp, options);
        using var stream = new MemoryStream();
        OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Bmp, stream, options, encoded.Length, CancellationToken.None);
        Assert.Equal(encoded, stream.ToArray());
        Assert.Equal(encoded.Length, BitConverter.ToInt32(encoded, 2));
        Assert.Equal(width, BitConverter.ToInt32(encoded, 18));
        Assert.Equal(height, BitConverter.ToInt32(encoded, 22));
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                int source = (y * width + x) * 4;
                int output = 122 + ((height - y - 1) * width + x) * 4;
                Assert.Equal(before[source + 2], encoded[output]);
                Assert.Equal(before[source + 1], encoded[output + 1]);
                Assert.Equal(before[source], encoded[output + 2]);
                Assert.Equal(before[source + 3], encoded[output + 3]);
            }
        }
        Assert.Equal(before, image.GetPixels());
        Assert.True(stream.CanWrite);
    }

    [Theory]
    [InlineData(0.01D, 1)]
    [InlineData(0.0127D, 1)]
    [InlineData(54546084.6592D, int.MaxValue)]
    [InlineData(double.MaxValue, int.MaxValue)]
    public void BmpDensityUsesPositiveSignedPixelsPerMeter(double dpi, int expected) {
        byte[] encoded = OfficeRasterImageEncoder.Encode(Pattern(2, 1), OfficeImageExportFormat.Bmp,
            new OfficeRasterEncodingOptions { DpiX = dpi, DpiY = dpi });
        Assert.Equal(expected, BitConverter.ToInt32(encoded, 38));
        Assert.Equal(expected, BitConverter.ToInt32(encoded, 42));
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(-1D)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void InvalidBmpDensityDoesNotMutateCallerStream(double invalid) {
        OfficeRasterImage image = Pattern(2, 1);
        byte[] before = image.GetPixels();
        using var stream = new MemoryStream();
        stream.WriteByte(0xAB);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterImageEncoder.EncodeTo(image,
            OfficeImageExportFormat.Bmp, stream, new OfficeRasterEncodingOptions { DpiX = 96D, DpiY = invalid }));
        Assert.Equal(new byte[] { 0xAB }, stream.ToArray());
        Assert.Equal(1L, stream.Position);
        Assert.Equal(before, image.GetPixels());
        Assert.True(stream.CanWrite);
    }

    [Fact]
    public void SuppressedBmpDensityDoesNotRequireDpiValues() {
        byte[] encoded = OfficeRasterImageEncoder.Encode(Pattern(1, 1), OfficeImageExportFormat.Bmp,
            new OfficeRasterEncodingOptions { WriteResolutionMetadata = false, DpiX = double.NaN, DpiY = -1D });
        Assert.Equal(0, BitConverter.ToInt32(encoded, 38));
        Assert.Equal(0, BitConverter.ToInt32(encoded, 42));
    }

    [Fact]
    public void SingleWideBmpRowUsesBoundedWritesAndCancelsBeforeItIsFullyProduced() {
        OfficeRasterImage image = Pattern(20_000, 1);
        byte[] before = image.GetPixels();
        using var cancellation = new CancellationTokenSource();
        using var stream = new ObservingStream(cancellation);
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageEncoder.EncodeTo(image,
            OfficeImageExportFormat.Bmp, stream, null, 100_000L, cancellation.Token));
        Assert.InRange(stream.LargestWrite, 1, 64 * 1024);
        Assert.InRange(stream.BytesWritten, 123L, before.Length);
        Assert.Equal(before, image.GetPixels());
        Assert.True(stream.CanWrite);
    }

    [Fact]
    public void BmpKnownOutputLimitIsRejectedBeforeMaterialization() {
        OfficeRasterImage image = Pattern(17, 3);
        byte[] before = image.GetPixels();
        var failure = Assert.Throws<OfficeImageExportBatchLimitException>(() =>
            OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Bmp, null, 128L));
        Assert.Equal(122L + before.Length, failure.Actual);
        Assert.Equal(128L, failure.Maximum);
        Assert.Equal(before, image.GetPixels());
    }

    [Theory]
    [InlineData(int.MaxValue, int.MaxValue)]
    [InlineData(50_000_000, 1)]
    public void BmpPlannerRejectsOversizedOutputWithoutAnImageAllocation(int width, int height) {
        Assert.Throws<ArgumentException>(() => OfficeBmpWriter.GetEncodedSize(width, height));
    }

    [Fact]
    public void BmpMaterializationIncludesRetainedExportInputsInWorkingSet() {
        OfficeRasterImage image = Pattern(1, 1);
        Assert.Throws<ArgumentException>(() => OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Bmp,
            null, 1024L, CancellationToken.None, OfficeRasterGuards.MaximumDecodedBytes - 130L));
    }

    [Fact]
    public void BmpWriteFailureLeavesSourceAndCallerOwnershipIntact() {
        OfficeRasterImage image = Pattern(17, 3);
        byte[] before = image.GetPixels();
        using var stream = new ObservingStream(null, failOnData: true);
        Assert.Throws<IOException>(() => OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Bmp, stream));
        Assert.Equal(122L, stream.BytesWritten);
        Assert.Equal(before, image.GetPixels());
        Assert.True(stream.CanWrite);
    }

    [Fact]
    public void IncrementalCapacityBoundIncludesNonPowerOfTwoSeedAndCoexistingArrays() {
        using var stream = new MemoryStream();
        long bound = OfficeRasterOutput.GetMemoryStreamWritePeakBytes(stream, 128_000, false);
        stream.Write(new byte[80_000], 0, 80_000);
        long oldCapacity = stream.Capacity;
        long singleBound = OfficeRasterOutput.GetMemoryStreamSingleWritePeakBytes(stream, 48_000, false);
        stream.Write(new byte[48_000], 0, 48_000);
        long observedGrowthBytes = oldCapacity + stream.Capacity + 48L;
        Assert.Equal(80_000L, oldCapacity);
        Assert.Equal(160_000, stream.Capacity);
        Assert.True(bound >= observedGrowthBytes);
        Assert.Equal(observedGrowthBytes, singleBound);
    }

    [Theory]
    [InlineData(0, 0, 0)]
    [InlineData(70_001, 70_000, 0)]
    [InlineData(100_000, 0, 90_000)]
    public void FixedBlockPlanCoversObservedGrowthAndMaterialization(int initialCapacity, int initialPosition, int initialLength) {
        using var stream = new MemoryStream(initialCapacity);
        stream.SetLength(initialLength);
        stream.Position = initialPosition;
        const int remaining = 100_003;
        long bound = OfficeRasterOutput.GetMemoryStreamBlockWritePeakBytes(stream, remaining, true, 16_384, 122, 16_384);
        long observed = stream.Capacity + 24L;
        var block = new byte[16_384];
        Append(122); Append(block.Length);
        for (int bytes = remaining - 122 - block.Length; bytes > 0; bytes -= block.Length) Append(Math.Min(bytes, block.Length));
        observed = Math.Max(observed, stream.Capacity + stream.Length + 48L);
        Assert.True(bound >= observed);
        void Append(int count) {
            int old = stream.Capacity;
            stream.Write(block, 0, count);
            observed = Math.Max(observed, old == stream.Capacity ? old + 24L : old + stream.Capacity + 48L);
        }
    }

    [Fact]
    public void NoGrowthPlanCountsEntireExposedBackingAndExistingLength() {
        var backing = new byte[10_000];
        using var stream = new MemoryStream(backing, 100, 500, writable: true, publiclyVisible: true);
        stream.Position = 20;
        Assert.Equal(10_024L, OfficeRasterOutput.GetMemoryStreamWritePeakBytes(stream, 100, false));
        Assert.Equal(10_548L, OfficeRasterOutput.GetMemoryStreamSingleWritePeakBytes(stream, 100, true));
    }

    [Fact]
    public void MetadataRewriteRejectsActualSecondResizeBeforeMutation() {
        using var stream = new OfficeMetadataRewriteStream(OfficeRasterGuards.MaximumDecodedBytes - 220_000L, 0, CancellationToken.None);
        stream.Write(new byte[80_000], 0, 80_000);
        Assert.Throws<ArgumentException>(() => stream.Write(new byte[48_000], 0, 48_000));
        Assert.Equal(80_000L, stream.Length);
        Assert.Equal(80_000, stream.Capacity);
        Assert.True(stream.CanWrite);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Bmp)]
    [InlineData(OfficeImageExportFormat.Pbm)]
    [InlineData(OfficeImageExportFormat.Tga)]
    public void StaticWritersObservePreCancellationWithoutMutatingCallerOutput(OfficeImageExportFormat format) {
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        using var stream = new MemoryStream(); stream.WriteByte(0xAB);
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageEncoder.EncodeTo(Pattern(2, 1), format,
            stream, null, 1024L, cancellation.Token));
        Assert.Equal(new byte[] { 0xAB }, stream.ToArray());
        Assert.True(stream.CanWrite);
    }

    [Fact]
    public void TgaBoundedWritesRetainExactAlphaAndStopWithinOneWideRow() {
        OfficeRasterImage image = Pattern(20_000, 1);
        byte[] before = image.GetPixels();
        byte[] encoded = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Tga);
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? decoded));
        Assert.Equal(before, decoded!.GetPixels());
        using var cancellation = new CancellationTokenSource();
        using var stream = new ObservingStream(cancellation, headerLength: 18);
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageEncoder.EncodeTo(image,
            OfficeImageExportFormat.Tga, stream, null, 100_000L, cancellation.Token));
        Assert.InRange(stream.LargestWrite, 1, 64 * 1024);
        Assert.InRange(stream.BytesWritten, 19L, before.Length);
        Assert.Equal(before, image.GetPixels());
        Assert.True(stream.CanWrite);
    }

    [Fact]
    public void TallNarrowPbmPadsEveryRowAndPreservesItsLastPixel() {
        const int height = 32_769;
        var image = new OfficeRasterImage(1, height, OfficeColor.White);
        image.SetPixel(0, height - 1, OfficeColor.Black);
        byte[] encoded = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Pbm);
        int headerLength = Encoding.ASCII.GetByteCount("P4\n1 " + height + "\n");
        Assert.Equal(headerLength + height, encoded.Length);
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? decoded));
        Assert.Equal(OfficeColor.White, decoded!.GetPixel(0, height - 2));
        Assert.Equal(OfficeColor.Black, decoded.GetPixel(0, height - 1));
    }

    private static OfficeRasterImage Pattern(int width, int height) {
        var image = new OfficeRasterImage(width, height);
        for (int at = 0; at < image.PixelBuffer.Length; at++) image.PixelBuffer[at] = (byte)(at * 73 + at / 7);
        return image;
    }

    private sealed class ObservingStream : Stream {
        private readonly CancellationTokenSource? _cancellation;
        private readonly bool _failOnData;
        private readonly int _headerLength;
        internal ObservingStream(CancellationTokenSource? cancellation, bool failOnData = false, int headerLength = 122) {
            _cancellation = cancellation; _failOnData = failOnData; _headerLength = headerLength;
        }
        internal int LargestWrite { get; private set; }
        internal long BytesWritten { get; private set; }
        public override void Write(byte[] buffer, int offset, int count) {
            LargestWrite = Math.Max(LargestWrite, count);
            if (BytesWritten >= _headerLength && _failOnData) throw new IOException("Synthetic caller write failure.");
            BytesWritten += count;
            if (BytesWritten > _headerLength) _cancellation?.Cancel();
        }
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}

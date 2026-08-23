using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Compares materialized output with caller-owned streaming across representative image shapes.</summary>
[MemoryDiagnoser]
public class ImageStreamingEncodeBenchmarks {
    private OfficeRasterImage _image = null!;
    private OfficeRasterEncodingOptions _options = null!;
    private CountingWriteStream _destination = null!;

    [ParamsSource(nameof(ScenarioIds))]
    public string ScenarioId { get; set; } = ImageBenchmarkScenarios.Tiny.Id;

    [Params(
        OfficeImageExportFormat.Png,
        OfficeImageExportFormat.Jpeg,
        OfficeImageExportFormat.Tiff,
        OfficeImageExportFormat.Webp)]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    public IEnumerable<string> ScenarioIds => ImageBenchmarkScenarios.TimedIds;

    [GlobalSetup]
    public void Setup() {
        ImageBenchmarkScenario scenario = ImageBenchmarkScenarios.Get(ScenarioId);
        _image = scenario.CreateImage();
        _options = CreateOptions();
        _destination = new CountingWriteStream();

        byte[] materialized = Materialized();
        using var streamed = new MemoryStream();
        OfficeRasterImageEncoder.EncodeTo(_image, Format, streamed, _options);
        byte[] streamedBytes = streamed.ToArray();
        if (materialized.Length != streamedBytes.Length && Format != OfficeImageExportFormat.Png) {
            throw new InvalidOperationException(
                $"{ScenarioId} {Format} stream length {streamedBytes.Length:N0} did not match materialized length {materialized.Length:N0}.");
        }

        OfficeRasterImage expected = ImageBenchmarkCorpus.Decode(materialized, ScenarioId + " materialized " + Format);
        OfficeRasterImage actual = ImageBenchmarkCorpus.Decode(streamedBytes, ScenarioId + " streamed " + Format);
        if (!string.Equals(
                ImageBenchmarkCorpus.PixelHash(expected),
                ImageBenchmarkCorpus.PixelHash(actual),
                StringComparison.Ordinal)) {
            throw new InvalidOperationException($"{ScenarioId} {Format} stream output did not preserve the materialized pixels.");
        }
    }

    [Benchmark(Baseline = true)]
    public byte[] Materialized() =>
        OfficeRasterImageEncoder.Encode(_image, Format, _options);

    [Benchmark]
    public long Streamed() {
        _destination.Reset();
        OfficeRasterImageEncoder.EncodeTo(_image, Format, _destination, _options);
        return _destination.BytesWritten;
    }

    private static OfficeRasterEncodingOptions CreateOptions() => new() {
        DpiX = 144D,
        DpiY = 120D,
        Png = new OfficePngEncodeOptions {
            Compression = OfficePngCompression.Optimal
        },
        Jpeg = new OfficeJpegEncodeOptions {
            Quality = 85,
            Subsampling = OfficeJpegSubsampling.Y420,
            Background = OfficeColor.White
        },
        Tiff = new OfficeTiffEncodeOptions {
            Compression = OfficeTiffCompression.PackBits
        }
    };

    private sealed class CountingWriteStream : Stream {
        internal long BytesWritten { get; private set; }

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => BytesWritten;

        public override long Position {
            get => BytesWritten;
            set => throw new NotSupportedException();
        }

        internal void Reset() => BytesWritten = 0;
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) {
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset > buffer.Length - count) {
                throw new ArgumentOutOfRangeException(nameof(count));
            }
            BytesWritten = checked(BytesWritten + count);
        }

        public override void WriteByte(byte value) =>
            BytesWritten = checked(BytesWritten + 1);
    }
}

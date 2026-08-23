using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Measures encoder time and managed allocation growth from identical deterministic RGBA sources.</summary>
[MemoryDiagnoser]
public class ImageEncodeBenchmarks {
    private OfficeRasterImage _image = null!;

    [Params(256, 512, 1024)]
    public int Size { get; set; } = 512;

    [GlobalSetup]
    public void Setup() {
        _image = ImageBenchmarkCorpus.CreatePattern(Size, Size);
        ValidateLossless(Png(), OfficeImageFormat.Png, nameof(Png));
        ValidateLossless(Tiff(), OfficeImageFormat.Tiff, nameof(Tiff));
        ValidateLossless(TiffDeflate(), OfficeImageFormat.Tiff, nameof(TiffDeflate));
        ValidateLossless(Webp(), OfficeImageFormat.Webp, nameof(Webp));
        ImageBenchmarkCorpus.AssertIdentified(Jpeg(), OfficeImageFormat.Jpeg, _image.Width, _image.Height, nameof(Jpeg));
        ImageBenchmarkCorpus.AssertIdentified(Jpeg420(), OfficeImageFormat.Jpeg, _image.Width, _image.Height, nameof(Jpeg420));
        ImageBenchmarkCorpus.AssertIdentified(
            Jpeg420ProgressiveOptimized(),
            OfficeImageFormat.Jpeg,
            _image.Width,
            _image.Height,
            nameof(Jpeg420ProgressiveOptimized));
    }

    [Benchmark(Baseline = true)]
    public byte[] Png() => OfficeRasterImageEncoder.Encode(_image, OfficeImageExportFormat.Png);

    [Benchmark]
    public byte[] Jpeg() => OfficeRasterImageEncoder.Encode(_image, OfficeImageExportFormat.Jpeg);

    [Benchmark]
    public byte[] Jpeg420() => OfficeJpegCodec.Encode(_image, new OfficeJpegEncodeOptions {
        Quality = 85,
        Subsampling = OfficeJpegSubsampling.Y420
    });

    [Benchmark]
    public byte[] Jpeg420ProgressiveOptimized() => OfficeJpegCodec.Encode(_image, new OfficeJpegEncodeOptions {
        Quality = 85,
        Subsampling = OfficeJpegSubsampling.Y420,
        Progressive = true,
        OptimizeHuffman = true
    });

    [Benchmark]
    public byte[] Tiff() => OfficeRasterImageEncoder.Encode(_image, OfficeImageExportFormat.Tiff);

    [Benchmark]
    public byte[] TiffDeflate() => OfficeTiffCodec.Encode(_image, new OfficeTiffEncodeOptions {
        Compression = OfficeTiffCompression.Deflate
    });

    [Benchmark]
    public byte[] Webp() => OfficeRasterImageEncoder.Encode(_image, OfficeImageExportFormat.Webp);

    private void ValidateLossless(byte[] encoded, OfficeImageFormat format, string operation) {
        ImageBenchmarkCorpus.AssertIdentified(encoded, format, _image.Width, _image.Height, operation);
        OfficeRasterImage decoded = ImageBenchmarkCorpus.Decode(encoded, operation);
        if (!string.Equals(ImageBenchmarkCorpus.PixelHash(decoded), ImageBenchmarkCorpus.PixelHash(_image), StringComparison.Ordinal)) {
            throw new InvalidOperationException(operation + " did not preserve the RGBA pixels.");
        }
    }
}

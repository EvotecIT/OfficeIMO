using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Measures encoders from one identical deterministic RGBA source.</summary>
[MemoryDiagnoser]
public class ImageEncodeBenchmarks {
    private OfficeRasterImage _image = null!;

    [GlobalSetup]
    public void Setup() {
        _image = ImageBenchmarkCorpus.CreatePattern();
        ValidateLossless(Png(), OfficeImageFormat.Png, nameof(Png));
        ValidateLossless(Tiff(), OfficeImageFormat.Tiff, nameof(Tiff));
        ValidateLossless(Webp(), OfficeImageFormat.Webp, nameof(Webp));
        ImageBenchmarkCorpus.AssertIdentified(Jpeg(), OfficeImageFormat.Jpeg, _image.Width, _image.Height, nameof(Jpeg));
    }

    [Benchmark(Baseline = true)]
    public byte[] Png() => OfficeRasterImageEncoder.Encode(_image, OfficeImageExportFormat.Png);

    [Benchmark]
    public byte[] Jpeg() => OfficeRasterImageEncoder.Encode(_image, OfficeImageExportFormat.Jpeg);

    [Benchmark]
    public byte[] Tiff() => OfficeRasterImageEncoder.Encode(_image, OfficeImageExportFormat.Tiff);

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

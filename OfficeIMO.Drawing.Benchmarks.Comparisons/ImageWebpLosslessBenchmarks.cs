using BenchmarkDotNet.Attributes;
using ImageMagick;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Equivalent static lossless WebP encode and decode comparisons.</summary>
[MemoryDiagnoser]
public class ImageWebpLosslessBenchmarks {
    private OfficeRasterImage _image = null!;
    private MagickImage _magick = null!;
    private byte[] _officeEncoded = null!;
    private byte[] _magickEncoded = null!;

    [GlobalSetup]
    public void Setup() {
        _image = ImageBenchmarkCorpus.CreatePattern(512, 512);
        _magick = ImageComparisonAdapters.CreateMagickImage(_image.GetPixels(), _image.Width, _image.Height);
        _officeEncoded = ImageComparisonAdapters.EncodeOfficeImoWebpLossless(_image);
        _magickEncoded = ImageComparisonAdapters.EncodeMagickWebpLossless(_magick);
        ImageComparisonValidation.ValidateLosslessInterchange("lossless WebP benchmark", _image, _officeEncoded, _magickEncoded, OfficeImageFormat.Webp);
    }

    [GlobalCleanup]
    public void Cleanup() => _magick.Dispose();

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMOEncode() => ImageComparisonAdapters.EncodeOfficeImoWebpLossless(_image);

    [Benchmark]
    public byte[] MagickNETEncode() => ImageComparisonAdapters.EncodeMagickWebpLossless(_magick);

    [Benchmark]
    public byte[] OfficeIMODecodeExternal() => ImageComparisonAdapters.DecodeOfficeImoRgba(_magickEncoded);

    [Benchmark]
    public byte[] MagickNETDecodeExternal() => ImageComparisonAdapters.DecodeMagickRgba(_magickEncoded);
}

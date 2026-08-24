using BenchmarkDotNet.Attributes;
using ImageMagick;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Equivalent single-frame, lossless TIFF LZW encode comparisons.</summary>
[MemoryDiagnoser]
public class ImageTiffLzwEncodeBenchmarks {
    private OfficeRasterImage _image = null!;
    private MagickImage _magick = null!;

    [GlobalSetup]
    public void Setup() {
        _image = ImageBenchmarkCorpus.CreatePattern(512, 512);
        _magick = ImageComparisonAdapters.CreateMagickImage(_image.GetPixels(), _image.Width, _image.Height);
        ImageComparisonAdapters.ConfigureMagickTiffLzw(_magick);
        byte[] officeEncoded = ImageComparisonAdapters.EncodeOfficeImoTiffLzw(_image);
        byte[] magickEncoded = ImageComparisonAdapters.EncodeMagickTiffLzw(_magick);
        ImageComparisonValidation.ValidateLosslessInterchange(
            "TIFF LZW encode benchmark", _image, officeEncoded, magickEncoded, OfficeImageFormat.Tiff);
    }

    [GlobalCleanup]
    public void Cleanup() => _magick.Dispose();

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMOEncode() => ImageComparisonAdapters.EncodeOfficeImoTiffLzw(_image);

    [Benchmark]
    public byte[] MagickNETEncode() => ImageComparisonAdapters.EncodeMagickTiffLzw(_magick);
}

/// <summary>Equivalent single-frame, lossless TIFF LZW decode comparisons.</summary>
[MemoryDiagnoser]
public class ImageTiffLzwDecodeBenchmarks {
    private byte[] _magickEncoded = null!;

    [GlobalSetup]
    public void Setup() {
        OfficeRasterImage image = ImageBenchmarkCorpus.CreatePattern(512, 512);
        using MagickImage magick = ImageComparisonAdapters.CreateMagickImage(
            image.GetPixels(), image.Width, image.Height);
        ImageComparisonAdapters.ConfigureMagickTiffLzw(magick);
        byte[] officeEncoded = ImageComparisonAdapters.EncodeOfficeImoTiffLzw(image);
        _magickEncoded = ImageComparisonAdapters.EncodeMagickTiffLzw(magick);
        ImageComparisonValidation.ValidateLosslessInterchange(
            "TIFF LZW decode benchmark", image, officeEncoded, _magickEncoded, OfficeImageFormat.Tiff);
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMODecodeExternal() => ImageComparisonAdapters.DecodeOfficeImoRgba(_magickEncoded);

    [Benchmark]
    public byte[] MagickNETDecodeExternal() => ImageComparisonAdapters.DecodeMagickRgba(_magickEncoded);
}

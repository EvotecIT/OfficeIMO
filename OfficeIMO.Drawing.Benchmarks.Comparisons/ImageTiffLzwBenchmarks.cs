using BenchmarkDotNet.Attributes;
using ImageMagick;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Equivalent single-frame, lossless TIFF LZW encode and decode comparisons.</summary>
[MemoryDiagnoser]
public class ImageTiffLzwBenchmarks {
    private OfficeRasterImage _image = null!;
    private MagickImage _magick = null!;
    private byte[] _officeEncoded = null!;
    private byte[] _magickEncoded = null!;

    [GlobalSetup]
    public void Setup() {
        _image = ImageBenchmarkCorpus.CreatePattern(512, 512);
        _magick = ImageComparisonAdapters.CreateMagickImage(_image.GetPixels(), _image.Width, _image.Height);
        _officeEncoded = ImageComparisonAdapters.EncodeOfficeImoTiffLzw(_image);
        _magickEncoded = ImageComparisonAdapters.EncodeMagickTiffLzw(_magick);
        ImageComparisonValidation.ValidateLosslessInterchange("TIFF LZW benchmark", _image, _officeEncoded, _magickEncoded, OfficeImageFormat.Tiff);
    }

    [GlobalCleanup]
    public void Cleanup() => _magick.Dispose();

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMOEncode() => ImageComparisonAdapters.EncodeOfficeImoTiffLzw(_image);

    [Benchmark]
    public byte[] MagickNETEncode() => ImageComparisonAdapters.EncodeMagickTiffLzw(_magick);

    [Benchmark]
    public byte[] OfficeIMODecodeExternal() => ImageComparisonAdapters.DecodeOfficeImoRgba(_magickEncoded);

    [Benchmark]
    public byte[] MagickNETDecodeOfficeIMO() => ImageComparisonAdapters.DecodeMagickRgba(_officeEncoded);
}

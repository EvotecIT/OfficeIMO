using BenchmarkDotNet.Attributes;
using ImageMagick;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Equivalent single-frame, lossless TIFF LZW encode comparisons.</summary>
// No [MemoryDiagnoser]: Magick.NET allocate in native memory, which BenchmarkDotNet does not measure.
public class ImageTiffLzwEncodeBenchmarks {
    private OfficeRasterImage _image = null!;
    private MagickImage _magick = null!;
#if NET10_0_OR_GREATER
    private global::OneImage.Image _oneImage = null!;
#endif

    [GlobalSetup]
    public void Setup() {
        _image = ImageBenchmarkCorpus.CreatePattern(512, 512);
        _magick = ImageComparisonAdapters.CreateMagickImage(_image.GetPixels(), _image.Width, _image.Height);
        ImageComparisonAdapters.ConfigureMagickTiffLzw(_magick);
        byte[] officeEncoded = ImageComparisonAdapters.EncodeOfficeImoTiffLzw(_image);
        byte[] magickEncoded = ImageComparisonAdapters.EncodeMagickTiffLzw(_magick);
#if NET10_0_OR_GREATER
        _oneImage = OneImageComparisonAdapters.CreateImage(_image.GetPixels(), _image.Width, _image.Height);
        ImageComparisonValidation.ValidateLosslessEncode(
            "TIFF LZW", _image, officeEncoded, magickEncoded, OfficeImageFormat.Tiff, OneImageEncode());
#else
        ImageComparisonValidation.ValidateLosslessEncode(
            "TIFF LZW", _image, officeEncoded, magickEncoded, OfficeImageFormat.Tiff);
#endif
    }

    [GlobalCleanup]
    public void Cleanup() => _magick.Dispose();

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Encode")]
    public byte[] OfficeIMOEncode() => ImageComparisonAdapters.EncodeOfficeImoTiffLzw(_image);

    [Benchmark]
    [BenchmarkCategory("library:Magick.NET", "case:Encode")]
    public byte[] MagickNETEncode() => ImageComparisonAdapters.EncodeMagickTiffLzw(_magick);
#if NET10_0_OR_GREATER

    [Benchmark]
    [BenchmarkCategory("library:OneImage", "case:Encode")]
    public byte[] OneImageEncode() => OneImageComparisonAdapters.EncodeTiffLzw(_oneImage);
#endif
}

/// <summary>Equivalent single-frame, lossless TIFF LZW decode comparisons.</summary>
// No [MemoryDiagnoser]: Magick.NET allocate in native memory, which BenchmarkDotNet does not measure.
public class ImageTiffLzwDecodeBenchmarks {
    private byte[] _magickEncoded = null!;

    [GlobalSetup]
    public void Setup() {
        OfficeRasterImage image = ImageBenchmarkCorpus.CreatePattern(512, 512);
        using MagickImage magick = ImageComparisonAdapters.CreateMagickImage(
            image.GetPixels(), image.Width, image.Height);
        ImageComparisonAdapters.ConfigureMagickTiffLzw(magick);
        _magickEncoded = ImageComparisonAdapters.EncodeMagickTiffLzw(magick);
        ImageComparisonValidation.ValidateLosslessDecode("TIFF LZW", image, _magickEncoded, OfficeImageFormat.Tiff);
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:DecodeExternal")]
    public byte[] OfficeIMODecodeExternal() => ImageComparisonAdapters.DecodeOfficeImoRgba(_magickEncoded);

    [Benchmark]
    [BenchmarkCategory("library:Magick.NET", "case:DecodeExternal")]
    public byte[] MagickNETDecodeExternal() => ImageComparisonAdapters.DecodeMagickRgba(_magickEncoded);
#if NET10_0_OR_GREATER

    [Benchmark]
    [BenchmarkCategory("library:OneImage", "case:DecodeExternal")]
    public byte[] OneImageDecodeExternal() => OneImageComparisonAdapters.DecodeRgba(_magickEncoded);
#endif
}

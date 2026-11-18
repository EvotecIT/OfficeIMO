using BenchmarkDotNet.Attributes;
using ImageMagick;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Equivalent static lossless WebP encode comparisons.</summary>
[MemoryDiagnoser]
public class ImageWebpLosslessEncodeBenchmarks {
    private OfficeRasterImage _image = null!;
    private MagickImage _magick = null!;
#if NET10_0_OR_GREATER
    private global::OneImage.Image _oneImage = null!;
#endif

    [GlobalSetup]
    public void Setup() {
        _image = ImageBenchmarkCorpus.CreatePattern(512, 512);
        _magick = ImageComparisonAdapters.CreateMagickImage(_image.GetPixels(), _image.Width, _image.Height);
        ImageComparisonAdapters.ConfigureMagickWebpLossless(_magick);
        byte[] officeEncoded = ImageComparisonAdapters.EncodeOfficeImoWebpLossless(_image);
        byte[] magickEncoded = ImageComparisonAdapters.EncodeMagickWebpLossless(_magick);
        ImageComparisonValidation.ValidateLosslessInterchange(
            "lossless WebP encode benchmark", _image, officeEncoded, magickEncoded, OfficeImageFormat.Webp);
#if NET10_0_OR_GREATER
        _oneImage = OneImageComparisonAdapters.CreateImage(_image.GetPixels(), _image.Width, _image.Height);
        ImageComparisonValidation.ValidateOneImageLosslessInterchange(
            "lossless WebP encode benchmark", _image, OneImageEncode(), officeEncoded, magickEncoded, OfficeImageFormat.Webp);
#endif
    }

    [GlobalCleanup]
    public void Cleanup() => _magick.Dispose();

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMOEncode() => ImageComparisonAdapters.EncodeOfficeImoWebpLossless(_image);

    [Benchmark]
    public byte[] MagickNETEncode() => ImageComparisonAdapters.EncodeMagickWebpLossless(_magick);
#if NET10_0_OR_GREATER

    [Benchmark]
    [BenchmarkCategory("library:oneimage")]
    public byte[] OneImageEncode() => OneImageComparisonAdapters.EncodeWebpLossless(_oneImage);
#endif
}

/// <summary>Equivalent static lossless WebP decode comparisons.</summary>
[MemoryDiagnoser]
public class ImageWebpLosslessDecodeBenchmarks {
    private byte[] _magickEncoded = null!;

    [GlobalSetup]
    public void Setup() {
        OfficeRasterImage image = ImageBenchmarkCorpus.CreatePattern(512, 512);
        using MagickImage magick = ImageComparisonAdapters.CreateMagickImage(
            image.GetPixels(), image.Width, image.Height);
        ImageComparisonAdapters.ConfigureMagickWebpLossless(magick);
        byte[] officeEncoded = ImageComparisonAdapters.EncodeOfficeImoWebpLossless(image);
        _magickEncoded = ImageComparisonAdapters.EncodeMagickWebpLossless(magick);
        ImageComparisonValidation.ValidateLosslessInterchange(
            "lossless WebP decode benchmark", image, officeEncoded, _magickEncoded, OfficeImageFormat.Webp);
#if NET10_0_OR_GREATER
        ImageComparisonValidation.ValidateOneImageLosslessInterchange(
            "lossless WebP decode benchmark",
            image,
            OneImageComparisonAdapters.EncodeWebpLossless(OneImageComparisonAdapters.CreateImage(image.GetPixels(), image.Width, image.Height)),
            officeEncoded,
            _magickEncoded,
            OfficeImageFormat.Webp);
#endif
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMODecodeExternal() => ImageComparisonAdapters.DecodeOfficeImoRgba(_magickEncoded);

    [Benchmark]
    public byte[] MagickNETDecodeExternal() => ImageComparisonAdapters.DecodeMagickRgba(_magickEncoded);
#if NET10_0_OR_GREATER

    [Benchmark]
    [BenchmarkCategory("library:oneimage")]
    public byte[] OneImageDecodeExternal() => OneImageComparisonAdapters.DecodeRgba(_magickEncoded);
#endif
}

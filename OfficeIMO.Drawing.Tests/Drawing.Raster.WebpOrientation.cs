using System;
using OfficeIMO.Drawing;
using Xunit;
namespace OfficeIMO.Tests;
public sealed class DrawingWebpOrientationTests {
    // Independent libwebp lossless payloads and Pillow ImageOps.exif_transpose pixel oracle.
    [Theory]
    [InlineData("UklGRmQAAABXRUJQVlA4WAoAAAAIAAAAAgAAAQAAVlA4TCQAAAAvAkAAAC8gEEjaH3qN+RcQFPk/2vwHH0QCg0CADPHiSET/IxZFWElGGgAAAE1NACoAAAAIAAEBEgADAAAAAQABAAAAAAAA", 3, 2, "/wAA/wD/AP8AAP////8A//8A//8A////")]
    [InlineData("UklGRmQAAABXRUJQVlA4WAoAAAAIAAAAAgAAAQAAVlA4TCQAAAAvAkAAAC8gEEjaH3qN+RcQFPk/2vwHH0QCg0CADPHiSET/IxZFWElGGgAAAE1NACoAAAAIAAEBEgADAAAAAQACAAAAAAAA", 3, 2, "AAD//wD/AP//AAD/AP////8A/////wD/")]
    [InlineData("UklGRmQAAABXRUJQVlA4WAoAAAAIAAAAAgAAAQAAVlA4TCQAAAAvAkAAAC8gEEjaH3qN+RcQFPk/2vwHH0QCg0CADPHiSET/IxZFWElGGgAAAE1NACoAAAAIAAEBEgADAAAAAQADAAAAAAAA", 3, 2, "AP////8A/////wD/AAD//wD/AP//AAD/")]
    [InlineData("UklGRmQAAABXRUJQVlA4WAoAAAAIAAAAAgAAAQAAVlA4TCQAAAAvAkAAAC8gEEjaH3qN+RcQFPk/2vwHH0QCg0CADPHiSET/IxZFWElGGgAAAE1NACoAAAAIAAEBEgADAAAAAQAEAAAAAAAA", 3, 2, "//8A//8A//8A/////wAA/wD/AP8AAP//")]
    [InlineData("UklGRmQAAABXRUJQVlA4WAoAAAAIAAAAAgAAAQAAVlA4TCQAAAAvAkAAAC8gEEjaH3qN+RcQFPk/2vwHH0QCg0CADPHiSET/IxZFWElGGgAAAE1NACoAAAAIAAEBEgADAAAAAQAFAAAAAAAA", 2, 3, "/wAA////AP8A/wD//wD//wAA//8A////")]
    [InlineData("UklGRmQAAABXRUJQVlA4WAoAAAAIAAAAAgAAAQAAVlA4TCQAAAAvAkAAAC8gEEjaH3qN+RcQFPk/2vwHH0QCg0CADPHiSET/IxZFWElGGgAAAE1NACoAAAAIAAEBEgADAAAAAQAGAAAAAAAA", 2, 3, "//8A//8AAP//AP//AP8A/wD///8AAP//")]
    [InlineData("UklGRmQAAABXRUJQVlA4WAoAAAAIAAAAAgAAAQAAVlA4TCQAAAAvAkAAAC8gEEjaH3qN+RcQFPk/2vwHH0QCg0CADPHiSET/IxZFWElGGgAAAE1NACoAAAAIAAEBEgADAAAAAQAHAAAAAAAA", 2, 3, "AP///wAA////AP//AP8A////AP//AAD/")]
    [InlineData("UklGRmQAAABXRUJQVlA4WAoAAAAIAAAAAgAAAQAAVlA4TCQAAAAvAkAAAC8gEEjaH3qN+RcQFPk/2vwHH0QCg0CADPHiSET/IxZFWElGGgAAAE1NACoAAAAIAAEBEgADAAAAAQAIAAAAAAAA", 2, 3, "AAD//wD///8A/wD//wD///8AAP///wD/")]
    public void PngConversionPreservesExifPresentation(string source, int width, int height, string expected) {
        Assert.True(OfficeImagePngConverter.TryConvertToPng(Convert.FromBase64String(source), out byte[] png));
        Assert.True(OfficeRasterImageDecoder.TryDecode(png, out var image));
        Assert.Equal(width, image!.Width);
        Assert.Equal(height, image.Height);
        Assert.Equal(Convert.FromBase64String(expected), image.GetPixels());
    }
    [Theory]
    [InlineData("UklGRpwAAABXRUJQVlA4WAoAAAAIAAAAAQAAAQAAVlA4TA8AAAAvAUAAAAcQ/Y/+ByKi/wEARVhJRmYAAABNTQAqAAAACAAGAQAABAAAAAEAAAACAQEABAAAAAEAAAACARIAAwAAAAEABgAAARoABQAAAAEAAABWARsABQAAAAEAAABeASgAAwAAAAEAAgAAAAAAAAAAASwAAAABAAAAlgAAAAE=", 2, 2)]
    [InlineData("UklGRpwAAABXRUJQVlA4WAoAAAAIAAAAAgAAAQAAVlA4TA8AAAAvAkAAAAcQ/Y/+ByKi/wEARVhJRmYAAABNTQAqAAAACAAGAQAABAAAAAEAAAADAQEABAAAAAEAAAACARIAAwAAAAEABgAAARoABQAAAAEAAABWARsABQAAAAEAAABeASgAAwAAAAEAAgAAAAAAAAAAASwAAAABAAAAlgAAAAE=", 2, 3)]
    public void OrientedPngKeepsPhysicalResolutionOnPresentationAxes(string source, int width, int height) {
        Assert.True(OfficeImagePngConverter.TryConvertToPng(Convert.FromBase64String(source), out byte[] png));
        Assert.True(OfficeImageReader.TryIdentify(png, null, out var info));
        Assert.Equal(width, info.Width);
        Assert.Equal(height, info.Height);
        Assert.InRange(info.DpiX, 149.9, 150.1);
        Assert.InRange(info.DpiY, 299.9, 300.1);
    }
}

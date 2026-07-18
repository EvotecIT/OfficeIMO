using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeImageExportBuilder_UsesOneFormatAndTerminalGrammar() {
        var options = new TestImageExportOptions();
        var builder = new TestImageExportBuilder(options);

        byte[] png = builder
            .AsPng()
            .WithScale(1.25D)
            .ToBytes();
        byte[] svg = builder
            .AsSvg()
            .WithBackground(OfficeColor.White)
            .ToBytes();

        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, png.Take(4).ToArray());
        Assert.Contains("<svg", System.Text.Encoding.UTF8.GetString(svg), StringComparison.Ordinal);
        Assert.Equal(1.25D, options.Scale);
        Assert.Equal(OfficeColor.White, options.BackgroundColor);
    }

    [Fact]
    public void OfficeImageExportBuilder_SavesTheConfiguredFormat() {
        var builder = new TestImageExportBuilder(new TestImageExportOptions());
        using var png = new MemoryStream();
        using var svg = new MemoryStream();

        OfficeImageExportResult pngResult = builder
            .AsPng()
            .Save(png);
        OfficeImageExportResult svgResult = builder.AsSvg().Save(svg);

        Assert.Equal(OfficeImageExportFormat.Png, pngResult.Format);
        Assert.Equal(OfficeImageExportFormat.Svg, svgResult.Format);
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, png.ToArray().Take(4).ToArray());
        Assert.Contains("<svg", System.Text.Encoding.UTF8.GetString(svg.ToArray()), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void OfficeImageExportBuilder_ExposesSharedRasterFormats(OfficeImageExportFormat format) {
        var options = new TestImageExportOptions();
        var builder = new TestImageExportBuilder(options);

        OfficeImageExportResult result = format switch {
            OfficeImageExportFormat.Jpeg => builder.AsJpeg().Export(),
            OfficeImageExportFormat.Tiff => builder.AsTiff().Export(),
            OfficeImageExportFormat.Webp => builder.AsWebp().Export(),
            _ => throw new ArgumentOutOfRangeException(nameof(format))
        };

        Assert.Equal(format, result.Format);
        Assert.Equal(format.GetMimeType(), OfficeImageReader.Identify(result.Bytes).MimeType);
    }

    [Fact]
    public void OfficeImageExportBuilder_ConfiguresRasterEncodingWithoutExposingOptions() {
        var options = new TestImageExportOptions();

        new TestImageExportBuilder(options)
            .WithRasterEncoding(settings => {
                settings.Jpeg.Quality = 73;
                settings.Tiff.Compression = OfficeTiffCompression.None;
            })
            .AsJpeg()
            .Export();

        Assert.Equal(73, options.RasterEncoding.Jpeg.Quality);
        Assert.Equal(OfficeTiffCompression.None, options.RasterEncoding.Tiff.Compression);
    }

    [Fact]
    public void OfficeImageExportBatchBuilder_SavesTheConfiguredFormat() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            var builder = new TestImageExportBatchBuilder(new TestImageExportOptions());

            IReadOnlyList<OfficeImageExportResult> pngResults = builder
                .AsPng()
                .Save(folder);
            IReadOnlyList<OfficeImageExportResult> svgResults = builder.AsSvg().Save(folder);

            Assert.Equal(OfficeImageExportFormat.Png, Assert.Single(pngResults).Format);
            Assert.Equal(OfficeImageExportFormat.Svg, Assert.Single(svgResults).Format);
            Assert.True(File.Exists(Path.Combine(folder, "batch.png")));
            Assert.True(File.Exists(Path.Combine(folder, "batch.svg")));
        } finally {
            if (Directory.Exists(folder)) {
                Directory.Delete(folder, recursive: true);
            }
        }
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Jpeg, "batch.jpg")]
    [InlineData(OfficeImageExportFormat.Tiff, "batch.tiff")]
    [InlineData(OfficeImageExportFormat.Webp, "batch.webp")]
    public void OfficeImageExportBatchBuilder_UsesTheConfiguredRasterExtension(
        OfficeImageExportFormat format,
        string expectedFileName) {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            var builder = new TestImageExportBatchBuilder(new TestImageExportOptions());

            builder.As(format).Save(folder);

            Assert.True(File.Exists(Path.Combine(folder, expectedFileName)));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void OfficeImageExportBatchBuilder_UsesPortableUniqueFileNames() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            var builder = new TestImageExportBatchBuilder(
                new TestImageExportOptions(),
                "Quarter:1",
                "Quarter?1",
                "CON",
                "COM¹",
                "LPT³.log",
                " trailing. ",
                "A/B");

            builder.Save(folder);

            Assert.True(File.Exists(Path.Combine(folder, "Quarter_1.png")));
            Assert.True(File.Exists(Path.Combine(folder, "Quarter_1-2.png")));
            Assert.True(File.Exists(Path.Combine(folder, "_CON.png")));
            Assert.True(File.Exists(Path.Combine(folder, "_COM¹.png")));
            Assert.True(File.Exists(Path.Combine(folder, "_LPT³.log.png")));
            Assert.True(File.Exists(Path.Combine(folder, "trailing.png")));
            Assert.True(File.Exists(Path.Combine(folder, "A_B.png")));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void OfficeImageExportBuilders_DoNotExposeDuplicateTerminalAliases() {
        string[] removedNames = {
            "AtScale", "OnBackground", "Preview", "HighResolution",
            "ExportPng", "ExportSvg", "ToPng", "ToPngBytes", "ToSvg", "ToSvgBytes",
            "SaveTo", "SaveAsPng", "SavePng", "SaveAsSvg", "SaveSvg"
        };
        string[] singleMethods = typeof(TestImageExportBuilder).GetMethods().Select(method => method.Name).ToArray();
        string[] batchMethods = typeof(TestImageExportBatchBuilder).GetMethods().Select(method => method.Name).ToArray();

        Assert.All(removedNames, name => Assert.DoesNotContain(name, singleMethods));
        Assert.All(removedNames, name => Assert.DoesNotContain(name, batchMethods));
    }
}

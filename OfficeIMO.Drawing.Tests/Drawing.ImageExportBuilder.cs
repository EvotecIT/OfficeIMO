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

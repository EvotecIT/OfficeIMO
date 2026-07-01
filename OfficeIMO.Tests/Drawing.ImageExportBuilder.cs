using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeImageExportBuilder_ExposesFriendlySingleExportAliases() {
        var options = new TestImageExportOptions();
        var builder = new TestImageExportBuilder(options);

        byte[] png = builder
            .AsSvg()
            .AtScale(1.25D)
            .ToPng();
        string svg = builder
            .WhiteBackground()
            .ToSvg();

        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, png.Take(4).ToArray());
        Assert.Contains("<svg", svg, StringComparison.Ordinal);
        Assert.Equal(1.25D, options.Scale);
        Assert.Equal(OfficeColor.White, options.BackgroundColor);
    }

    [Fact]
    public void OfficeImageExportBuilder_ExposesFriendlySaveAliases() {
        var builder = new TestImageExportBuilder(new TestImageExportOptions());
        using var png = new MemoryStream();
        using var svg = new MemoryStream();

        OfficeImageExportResult pngResult = builder
            .AsSvg()
            .SavePng(png);
        OfficeImageExportResult svgResult = builder
            .SaveSvg(svg);

        Assert.Equal(OfficeImageExportFormat.Png, pngResult.Format);
        Assert.Equal(OfficeImageExportFormat.Svg, svgResult.Format);
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, png.ToArray().Take(4).ToArray());
        Assert.Contains("<svg", System.Text.Encoding.UTF8.GetString(svg.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeImageExportBatchBuilder_ExposesFriendlySaveAliases() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            var builder = new TestImageExportBatchBuilder(new TestImageExportOptions());

            IReadOnlyList<OfficeImageExportResult> pngResults = builder
                .AsSvg()
                .SavePng(folder);
            IReadOnlyList<OfficeImageExportResult> svgResults = builder.SaveSvg(folder);

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
}

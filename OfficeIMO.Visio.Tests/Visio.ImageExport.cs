using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioImageExport {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void PageExport_EncodesEverySharedImageFormat(OfficeImageExportFormat format) {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Formats").Size(2, 1);
        page.AddRectangle(1, 0.5, 1.2, 0.5, "Format marker");
        var options = new VisioImageExportOptions { Scale = 0.5D, Supersampling = 1 };

        OfficeImageExportResult result = page.ExportImage(format, options);

        Assert.Equal(format, result.Format);
        Assert.Equal(96, result.Width);
        Assert.Equal(48, result.Height);
        if (format == OfficeImageExportFormat.Svg) {
            Assert.Contains("<svg", Encoding.UTF8.GetString(result.Bytes), StringComparison.Ordinal);
        } else {
            Assert.Equal(format.GetMimeType(), OfficeImageReader.Identify(result.Bytes).MimeType);
        }
    }

    [Fact]
    public void DocumentFluentBatchExport_SelectsPagesAndSavesPortableResults() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        document.AddPage("First").Size(2, 1).AddRectangle(1, 0.5, 1, 0.5, "One");
        document.AddPage("Second").Size(2, 1).AddEllipse(1, 0.5, 1, 0.5, "Two");
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            IReadOnlyList<OfficeImageExportResult> results = document.ToImages()
                .FromPage(1)
                .TakePages(1)
                .WithDpi(48D)
                .AsWebp()
                .Save(folder);

            OfficeImageExportResult result = Assert.Single(results);
            Assert.Equal("Second", result.Name);
            Assert.Equal("Visio page 2", result.Source);
            Assert.Equal("image/webp", OfficeImageReader.Identify(result.Bytes).MimeType);
            Assert.True(File.Exists(Path.Combine(folder, "Second.webp")));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void DocumentAndPageConvenienceMethods_ProduceRequestedRasterFormats() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Convenience").Size(2, 1);
        page.AddDiamond(1, 0.5, 1, 0.6, "Choice");
        var options = new VisioImageExportOptions { Scale = 0.5D, Supersampling = 1 };

        Assert.Equal("image/jpeg", OfficeImageReader.Identify(page.ToJpeg(options)).MimeType);
        Assert.Equal("image/tiff", OfficeImageReader.Identify(document.ToTiff(options)).MimeType);
        Assert.Equal("image/webp", OfficeImageReader.Identify(page.ToWebp(options)).MimeType);
        Assert.Equal(OfficeImageExportFormat.Svg, document.ToImage().WithDpi(48D).AsSvg().Export().Format);
    }

    [Fact]
    public void RasterExport_ReducesScaleWithAVisibleDiagnosticInsteadOfOverAllocating() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Large").Size(100, 100);
        page.AddRectangle(50, 50, 80, 80, "Large");
        var options = new VisioImageExportOptions {
            Scale = 1D,
            Supersampling = 1,
            MaximumRasterPixels = 10_000L
        };

        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png, options);

        Assert.True((long)result.Width * result.Height <= options.MaximumRasterPixels);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.RasterScaleReduced);
    }

    [Fact]
    public void ImageExportOptions_CloneRasterSettingsAcrossDocumentSelection() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        document.AddPage("Options").Size(2, 1).AddRectangle(1, 0.5, 1, 0.5, "Options");
        var options = new VisioImageExportOptions { Scale = 0.5D, Supersampling = 1 };
        options.RasterEncoding.Jpeg.Quality = 67;
        options.RasterEncoding.Jpeg.Progressive = true;

        OfficeImageExportResult result = document.ExportImage(OfficeImageExportFormat.Jpeg, options);

        Assert.Equal("image/jpeg", OfficeImageReader.Identify(result.Bytes).MimeType);
        Assert.True(result.Bytes.Contains((byte)0xC2));
    }

    [Fact]
    public void SvgExport_RejectsOverflowingScaleBeforeRendering() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Overflow").Size(2, 1);
        var options = new VisioImageExportOptions { Scale = double.MaxValue };

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            page.ExportImage(OfficeImageExportFormat.Svg, options));
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg)]
    public void PackagePreview_UsesCallerCodecThroughCanonicalExport(OfficeImageExportFormat format) {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("CallerCodec").Size(2, 1);
        AddCustomPreviewShape(page);
        var codec = new SolidImageCodec(OfficeColor.FromRgb(12, 90, 180));
        var options = new VisioImageExportOptions {
            Scale = 0.5D,
            Supersampling = 1,
            ImageCodec = codec
        };

        OfficeImageExportResult result = page.ExportImage(format, options);

        Assert.Equal(1, codec.DecodeCalls);
        Assert.Contains(
            result.Diagnostics,
            diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec);
        Assert.DoesNotContain(
            result.Diagnostics,
            diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
        if (format == OfficeImageExportFormat.Svg) {
            Assert.Contains("data:image/png;base64,", Encoding.UTF8.GetString(result.Bytes), StringComparison.Ordinal);
        }
    }

    private static void AddCustomPreviewShape(VisioPage page) {
        VisioMaster master = new("custom-preview", "CustomPreview", new VisioShape("master-shape", 0, 0, 1, 1, string.Empty));
        master.RawMasterRelationships.Add(new VisioAssets.MasterRelationshipContent {
            Id = "rIdImage",
            Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            Target = "../media/preview.custom",
            ContentType = "image/x-officeimo-test",
            Extension = ".custom",
            Data = new byte[] { 1, 2, 3, 4 }
        });
        VisioShape shape = page.AddRectangle(1, 0.5, 1, 0.7, string.Empty);
        shape.Master = master;
        shape.NameU = master.NameU;
        shape.SetUserCell("OfficeIMO.StencilPreviewImageRelationshipId", "rIdImage", "STR");
        shape.SetUserCell("OfficeIMO.StencilPreviewImageTarget", "../media/preview.custom", "STR");
    }

    private sealed class SolidImageCodec : IOfficeRasterImageCodec {
        private readonly OfficeColor _color;

        internal SolidImageCodec(OfficeColor color) {
            _color = color;
        }

        internal int DecodeCalls { get; private set; }

        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            DecodeCalls++;
            image = new OfficeRasterImage(2, 2, _color);
            return true;
        }
    }
}

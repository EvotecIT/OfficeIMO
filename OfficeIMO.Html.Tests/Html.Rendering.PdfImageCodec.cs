using System;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    // Independent two-frame WebP keeps this proof at the explicit caller-codec boundary.
    private const string CodecAnimatedImage = "UklGRogAAABXRUJQVlA4WAoAAAACAAAADwAADwAAQU5JTQYAAAAAAAAAAABBTk1GKgAAAAAAAAAAAA8AAA8AAGQAAAJWUDhMEQAAAC8PwAMAB1CoohSv/4GI6H8AAEFOTUYqAAAAAAAAAAAADwAADwAAZAAAAFZQOEwRAAAALw/AAwAHUKjiFaX/gYjofwAA";

    [Theory]
    [InlineData("")]
    [InlineData("opacity:.5")]
    [InlineData("clip-path:inset(0)")]
    public void HtmlPdf_ConfiguredRasterCodecPreservesNestedImages(string style) {
        var codec = new PdfRasterCodec(OfficeColor.Red);
        var document = HtmlConversionDocument.Parse("<img style='" + style + "' src='data:image/webp;base64," + CodecAnimatedImage + "'>");
        var result = document.RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Pdf, new HtmlToPdfOptions { ImageCodec = codec }));
        Assert.True(codec.Calls > 0);
        Assert.DoesNotContain(result.Output.Warnings, x => x.Code == HtmlPdfDiagnosticCodes.ImagePayloadOmitted);
        Assert.Equal(OfficeColor.Red, ExtractImagePixel(result.ToBytes()));
    }

    [Fact]
    public void HtmlPdf_ImagePreparationDoesNotReuseAnotherExportsCodecDecision() {
        var document = HtmlConversionDocument.Parse("<img src='data:image/webp;base64," + CodecAnimatedImage + "'>");
        byte[] Render(IOfficeRasterImageCodec? codec) => document.RenderToPdfResult(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions { ImageCodec = codec })).ToBytes();
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(Render(null)));
        Assert.Equal(OfficeColor.Red, ExtractImagePixel(Render(new PdfRasterCodec(OfficeColor.Red))));
        Assert.Equal(OfficeColor.Blue, ExtractImagePixel(Render(new PdfRasterCodec(OfficeColor.Blue))));
    }

    [Fact]
    public void HtmlPdf_ImageCodecCannotBypassRasterPixelLimit() {
        var codec = new PdfRasterCodec(OfficeColor.Red);
        var document = HtmlConversionDocument.Parse("<img src='data:image/webp;base64," + CodecAnimatedImage + "'>");
        var result = document.RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Pdf, new HtmlToPdfOptions { ImageCodec = codec, MaximumRasterPixels = 255 }));
        Assert.Equal(0, codec.Calls);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(result.ToBytes()));
        Assert.Contains(result.Output.Warnings, x => x.Code == HtmlPdfDiagnosticCodes.ImagePayloadOmitted);
    }

    [Fact]
    public void HtmlPdf_StaticImageSelectionReportsDiscardedAnimation() {
        // Independently generated one-frame GIF with a nonzero playback duration.
        var document = HtmlConversionDocument.Parse("<img src='data:image/gif;base64,R0lGODlhAQABAIEAAP8AAAAAAAAAAAAAACH/C05FVFNDQVBFMi4wAwEAAAAh+QQACgAAACwAAAAAAQABAAAIBAABBAQAOw=='>");
        var result = document.RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf));
        Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(result.ToBytes()));
        Assert.Contains(result.Output.Warnings, x => x.Code == HtmlPdfDiagnosticCodes.ImageStaticFrameSelected
            && x.LossKind == OfficeConversionLossKind.Omission);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPdf_StaticApngSelectionReportsLossAcrossDrawingEffects(bool drawingEffect) {
        const string apng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACGFjVEwAAAACAAAAAPONk3AAAAAaZmNUTAAAAAAAAAABAAAAAQAAAAAAAAAAAGQD6AAAs35jzQAAAA1JREFUeJxj+M/A8B8ABQAB/4mZPR0AAAAaZmNUTAAAAAEAAAABAAAAAQAAAAAAAAAAAGQD6AAAKA2JGQAAABFmZEFUAAAAAnicY2Bg+P8fAAMCAf/1e6XXAAAAAElFTkSuQmCC";
        var document = HtmlConversionDocument.Parse(ImageHtml("image/png", apng, drawingEffect));
        var result = document.RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf));
        Assert.Contains(result.Output.Warnings, x => x.Code == HtmlPdfDiagnosticCodes.ImageStaticFrameSelected
            && x.LossKind == OfficeConversionLossKind.Omission);
        Assert.NotEmpty(PdfCore.PdfImageExtractor.ExtractImages(result.ToBytes()));
        if (drawingEffect) Assert.Contains(result.Output.Warnings, x => x.Code == "HtmlPdfDrawingEffectRasterized");
    }

    [Theory]
    [InlineData(16, false, true)]
    [InlineData(1, false, false)]
    [InlineData(16, true, false)]
    public void HtmlPdf_RasterizedDrawingEffectUsesInspectedCodecBoundary(int size, bool malformed, bool visible) {
        byte[] bytes = Convert.FromBase64String(CodecAnimatedImage);
        if (malformed) { bytes[20] = 0; bytes[21] = 0; bytes[22] = 0; }
        var codec = new PdfRasterCodec(OfficeColor.Red, size);
        var document = HtmlConversionDocument.Parse(ImageHtml("image/webp", Convert.ToBase64String(bytes), true));
        var result = document.RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Pdf, new HtmlToPdfOptions { ImageCodec = codec }));
        Assert.Equal(malformed ? 0 : 1, codec.Calls);
        Assert.Contains(result.Output.Warnings, x => x.Code == "HtmlPdfDrawingEffectRasterized");
        // Malformed metadata is rejected during SVG import; an invalid caller
        // result is rejected later during raster preparation. Both report loss.
        if (!visible) Assert.Contains(result.Output.Warnings, x => x.Code ==
            (malformed ? HtmlRenderDiagnosticCodes.SvgContentUnsupported : HtmlPdfDiagnosticCodes.ImagePayloadOmitted)
            && x.LossKind == OfficeConversionLossKind.Omission);
        var embedded = Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(result.ToBytes()));
        Assert.True(OfficePngReader.TryDecode(embedded.Bytes, out var image));
        Assert.Equal(visible ? OfficeColor.Red : OfficeColor.Transparent, image!.GetPixel(8, 8));
    }

    private static string ImageHtml(string contentType, string base64, bool drawingEffect) {
        string src = "data:" + contentType + ";base64," + base64;
        if (!drawingEffect) return "<img src='" + src + "'>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16'>"
            + "<g style='mix-blend-mode:multiply'><image href='" + src + "' width='16' height='16'/></g></svg>";
        return "<img width='16' height='16' src='data:image/svg+xml;base64,"
            + Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)) + "'>";
    }

    private static OfficeColor ExtractImagePixel(byte[] pdf) {
        var embedded = Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(pdf));
        Assert.True(OfficePngReader.TryDecode(embedded.Bytes, out var image));
        return image!.GetPixel(8, 8);
    }

    private sealed class PdfRasterCodec : IOfficeRasterImageCodec {
        private readonly OfficeColor _color;
        internal int Calls;
        private readonly int _size;
        internal PdfRasterCodec(OfficeColor color, int size = 16) { _color = color; _size = size; }
        public bool TryDecode(byte[] bytes, string? contentType, out OfficeRasterImage? image) {
            image = null;
            // This fixture codec supports WebP, not the containing SVG.
            if (bytes.Length < 12 || bytes[0] != (byte)'R' || bytes[8] != (byte)'W') return false;
            Calls++;
            image = new OfficeRasterImage(_size, _size, _color);
            return true;
        }
    }
}

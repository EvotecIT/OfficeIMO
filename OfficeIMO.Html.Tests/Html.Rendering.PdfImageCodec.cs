using System;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private const string CodecVp8Image = "UklGRjwAAABXRUJQVlA4IDAAAADQAQCdASoQABAAAUAmJaACdLoB+AADsAD+8ut//NgVzXPv9//S4P0uD9Lg/9KQAAA=";

    [Theory]
    [InlineData("")]
    [InlineData("opacity:.5")]
    [InlineData("clip-path:inset(0)")]
    public void HtmlPdf_ConfiguredRasterCodecPreservesNestedImages(string style) {
        var codec = new PdfRasterCodec(OfficeColor.Red);
        var document = HtmlConversionDocument.Parse("<img style='" + style + "' src='data:image/webp;base64," + CodecVp8Image + "'>");
        var result = document.RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Pdf, new HtmlToPdfOptions { ImageCodec = codec }));
        Assert.True(codec.Calls > 0);
        Assert.DoesNotContain(result.Output.Warnings, x => x.Code == HtmlPdfDiagnosticCodes.ImagePayloadOmitted);
        Assert.Equal(OfficeColor.Red, ExtractImagePixel(result.ToBytes()));
    }

    [Fact]
    public void HtmlPdf_ImagePreparationDoesNotReuseAnotherExportsCodecDecision() {
        var document = HtmlConversionDocument.Parse("<img src='data:image/webp;base64," + CodecVp8Image + "'>");
        byte[] Render(IOfficeRasterImageCodec? codec) => document.RenderToPdfResult(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions { ImageCodec = codec })).ToBytes();
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(Render(null)));
        Assert.Equal(OfficeColor.Red, ExtractImagePixel(Render(new PdfRasterCodec(OfficeColor.Red))));
        Assert.Equal(OfficeColor.Blue, ExtractImagePixel(Render(new PdfRasterCodec(OfficeColor.Blue))));
    }

    [Fact]
    public void HtmlPdf_ImageCodecCannotBypassRasterPixelLimit() {
        var codec = new PdfRasterCodec(OfficeColor.Red);
        var document = HtmlConversionDocument.Parse("<img src='data:image/webp;base64," + CodecVp8Image + "'>");
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

    private static OfficeColor ExtractImagePixel(byte[] pdf) {
        var embedded = Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(pdf));
        Assert.True(OfficePngReader.TryDecode(embedded.Bytes, out var image));
        return image!.GetPixel(8, 8);
    }

    private sealed class PdfRasterCodec : IOfficeRasterImageCodec {
        private readonly OfficeColor _color;
        internal int Calls;
        internal PdfRasterCodec(OfficeColor color) => _color = color;
        public bool TryDecode(byte[] bytes, string? contentType, out OfficeRasterImage? image) {
            Calls++;
            image = new OfficeRasterImage(16, 16, _color);
            return true;
        }
    }
}

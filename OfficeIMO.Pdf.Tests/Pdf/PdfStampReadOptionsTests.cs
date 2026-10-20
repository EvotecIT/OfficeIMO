using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfStampReadOptionsTests {
    [Theory]
    [InlineData("Text")]
    [InlineData("TextWatermark")]
    [InlineData("Image")]
    [InlineData("ImageWatermark")]
    [InlineData("OverlayPage")]
    [InlineData("UnderlayPage")]
    [InlineData("Content")]
    [InlineData("Watermark")]
    public void ExplicitReadOptionsRemainOnStampResultAndSubsequentMutations(string operation) {
        byte[] original = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Original body")).ToBytes();
        var originalOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxJavaScripts = 19 } };
        var explicitOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxJavaScripts = 3 } };
        PdfDocument source = PdfDocument.Load(original, originalOptions);
        byte[] image = PdfPngTestImages.CreateRgbPng(20, 80, 180);
        byte[] overlay = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Imported page")).ToBytes();

        PdfDocument stamped = operation switch {
            "Text" => source.Stamp.Text("Stamp", null, explicitOptions),
            "TextWatermark" => source.Stamp.TextWatermark("Watermark", null, explicitOptions),
            "Image" => source.Stamp.Image(image, new PdfImageStampOptions { Width = 24D, Height = 24D }, explicitOptions),
            "ImageWatermark" => source.Stamp.ImageWatermark(image, new PdfImageStampOptions { Width = 24D, Height = 24D }, explicitOptions),
            "OverlayPage" => source.Stamp.OverlayPage(overlay, null, explicitOptions),
            "UnderlayPage" => source.Stamp.UnderlayPage(overlay, null, explicitOptions),
            "Content" => source.Stamp.Content((canvas, _) => canvas.Text("Canvas", 30D, 30D, 200D, 24D), readOptions: explicitOptions),
            "Watermark" => source.Stamp.Watermark(new PdfWatermarkOptions { Text = "Watermark" }, explicitOptions),
            _ => throw new ArgumentOutOfRangeException(nameof(operation))
        };

        Assert.Equal(3, stamped.ReadOptions.Limits.MaxJavaScripts);
        Assert.Equal(19, source.ReadOptions.Limits.MaxJavaScripts);
        Assert.Equal(original, source.ToBytes());
        Assert.False(original.SequenceEqual(stamped.ToBytes()));
        Assert.Contains("Original body", stamped.Reader.Text(), StringComparison.Ordinal);

        PdfDocument subsequent = stamped.Stamp.Text("Subsequent stamp");
        Assert.Equal(3, subsequent.ReadOptions.Limits.MaxJavaScripts);
        Assert.Contains("Subsequent stamp", subsequent.Reader.Text(), StringComparison.Ordinal);
    }
}

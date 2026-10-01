using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPictureLayoutRegressionTests {
    private static RtfDocument CreatePicture(bool padding = false) {
        var raster = new OfficeRasterImage(100, 60, OfficeColor.Red);
        RtfDocument document = RtfDocument.Create();
        RtfImage image = document.AddImage(RtfImageFormat.Png, OfficePngWriter.Encode(raster));
        image.SourceWidth = 100;
        image.SourceHeight = 60;
        image.DesiredWidthTwips = 2000;
        image.DesiredHeightTwips = 1200;
        image.ScaleXPercent = 150;
        image.ScaleYPercent = 200;
        image.CropLeftTwips = padding ? -200 : 200;
        image.CropRightTwips = 400;
        image.CropTopTwips = 100;
        image.CropBottomTwips = 100;
        return document;
    }

    [Fact]
    public void Standalone_Pictures_Keep_Their_Paragraph_Boundaries_And_Inline_Pictures_Keep_Their_Run() {
        RtfDocument document = CreatePicture();
        document.InsertParagraph(0, "Before");
        document.AddParagraph("After");
        RtfParagraph inline = document.AddParagraph("Inline ");
        inline.AddImage(RtfImageFormat.Png, Assert.IsType<RtfImage>(document.Blocks[1]).Data);
        inline.AddText(" tail");
        string native = document.ToRtf();
        Assert.Contains("}\\par", native, StringComparison.Ordinal);
        RtfDocument reopened = RtfDocument.Read(native).Document;
        Assert.Collection(reopened.Blocks,
            block => Assert.Equal("Before", Assert.IsType<RtfParagraph>(block).ToPlainText()),
            block => Assert.IsType<RtfImage>(block),
            block => Assert.Equal("After", Assert.IsType<RtfParagraph>(block).ToPlainText()),
            block => {
                RtfParagraph paragraph = Assert.IsType<RtfParagraph>(block);
                Assert.Single(paragraph.Inlines.OfType<RtfImage>());
                Assert.Equal("Inline  tail", paragraph.ToPlainText());
            });
    }

    [Fact]
    public void Native_Picture_Transforms_Survive_Normalization_Cloning_And_Image_Replacement() {
        RtfDocument document = CreatePicture();
        RtfImage source = Assert.IsType<RtfImage>(document.Blocks[0]);
        RtfImage reopened = Assert.IsType<RtfImage>(RtfDocument.Read(document.ToRtf()).Document.Blocks[0]);
        Assert.Equal(150, reopened.ScaleXPercent);
        Assert.Equal(200, reopened.ScaleYPercent);
        Assert.Equal(200, reopened.CropLeftTwips);
        Assert.Equal(400, reopened.CropRightTwips);
        Assert.Equal(100, reopened.CropTopTwips);
        Assert.Equal(100, reopened.CropBottomTwips);
        Assert.Equal(2100, reopened.ResolveLayout().VisibleWidthTwips);
        Assert.Equal(2000, reopened.ResolveLayout().VisibleHeightTwips);
        RtfImage clone = Assert.IsType<RtfImage>(document.Clone().Blocks[0]);
        clone.CropLeftTwips = 0;
        Assert.Equal(200, source.CropLeftTwips);
        RtfLosslessEditor editor = RtfDocument.Read(document.ToRtf()).EditLossless();
        Assert.True(editor.ReplaceImage(0, source));
        Assert.Equal(2100, Assert.IsType<RtfImage>(editor.ToReadResult().Document.Blocks[0]).ResolveLayout().VisibleWidthTwips);
    }

    [Theory]
    [InlineData(false, 2100d)]
    [InlineData(true, 2700d)]
    public void Html_And_Saved_Word_Preserve_Visible_Crop_Geometry(bool padding, double visibleWidth) {
        RtfDocument document = CreatePicture(padding);
        string html = document.ToHtml(RtfToHtmlOptions.CreateRoundTripProfile());
        AngleSharp.Dom.IElement rendered = Assert.Single(new AngleSharp.Html.Parser.HtmlParser().ParseDocument(html).QuerySelectorAll("img"));
        Assert.Contains("overflow:hidden", rendered.ParentElement!.GetAttribute("style"), StringComparison.Ordinal);
        RtfImage htmlImage = Assert.Single(HtmlConversionDocument.Parse(html).ToRtfDocument().Paragraphs.SelectMany(paragraph => paragraph.Inlines).OfType<RtfImage>());
        Assert.Equal(visibleWidth, htmlImage.ResolveLayout().VisibleWidthTwips);
        Assert.Equal(padding ? -200 : 200, htmlImage.CropLeftTwips);
        using WordDocument word = document.ToWordDocument();
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(word._wordprocessingDocument));
        WordImage output = Assert.Single(word.Images);
        Assert.Equal(visibleWidth / 15d, output.Width);
        Assert.Equal(padding ? -10000 : 10000, output.CropLeft);
        Assert.Equal(20000, output.CropRight);
        using var bytes = new MemoryStream();
        word.Save(bytes);
        bytes.Position = 0;
        using WordDocument opened = WordDocument.Load(bytes);
        RtfImage native = Assert.IsType<RtfImage>(opened.ToRtfDocument().Blocks[0]);
        Assert.Equal(visibleWidth, native.ResolveLayout().VisibleWidthTwips);
        Assert.Equal(2000, native.ResolveLayout().VisibleHeightTwips);
        Assert.Equal(100, native.SourceWidth);
        Assert.Equal(60, native.SourceHeight);
    }

    [Fact]
    public void Pdf_Handles_Cropping_And_Padding_And_Rejects_Empty_Visible_Areas() {
        foreach (bool padding in new[] { false, true }) {
            RtfDocument document = CreatePicture(padding);
            var result = document.ToPdfDocumentResult();
            Assert.DoesNotContain(result.Warnings, warning => warning.Code is "ImageCropPaddingFlattened" or "ImageCropDecodeFailed" or "ImageLayoutInvalid");
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(result.Value.ToBytes());
            Assert.Single(pdf.GetPage(1).GetImages());
            if (padding) {
                var extracted = Assert.Single(OfficeIMO.Pdf.PdfReadDocument.Open(result.Value.ToBytes()).ExtractImages());
                Assert.Equal(90, extracted.Width);
                Assert.Equal(50, extracted.Height);
                Assert.True(OfficeRasterImageDecoder.TryDecode(extracted.Bytes, out OfficeRasterImage? raster));
                Assert.Equal(0, raster!.GetPixel(0, 25).A);
                Assert.Equal(OfficeColor.Red, raster.GetPixel(20, 25));
            }
        }
        RtfDocument invalid = CreatePicture();
        Assert.IsType<RtfImage>(invalid.Blocks[0]).CropLeftTwips = 2000;
        Assert.Contains(invalid.ToHtmlResult(RtfToHtmlOptions.CreateRoundTripProfile()).RtfReport.Diagnostics, item => item.Code == "RtfHtmlImageLayoutInvalid" && item.Action == RtfConversionAction.Blocked);
        Assert.Contains(invalid.ToPdfDocumentResult().Warnings, item => item.Code == "ImageLayoutInvalid");
        Assert.Contains(invalid.ToWordDocumentResult().Report.Diagnostics, item => item.Code == "RtfWordImagesOmitted" && item.Action == RtfConversionAction.Omitted);
    }
}

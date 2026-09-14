using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReverseImagePlacementSafetyTests {
    private static readonly byte[] Png = PdfPngTestImages.CreateRgbPng(8, 4);

    [Fact]
    public void InvisibleImagePlacementsAreSuppressedAcrossEditableAdapters() {
        byte[] source = CreateDocument()
            .Canvas(canvas => canvas.Effect(
                OfficeIMO.Drawing.OfficeTransform.Identity,
                0D,
                effect => effect.Image(Png, 20D, 30D, 80D, 40D)))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        Assert.Equal(0D, Assert.Single(Assert.Single(logical.Pages).Images).Placements[0].Opacity);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Empty(word.Value.Images);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfInvisibleImagePlacementSuppressed" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.DoesNotContain("data:image/", html.Value, StringComparison.Ordinal);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "InvisibleImagePlacementSuppressed" &&
            warning.LossKind == OfficeConversionLossKind.None);

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Empty(powerPoint.Value.Slides.SelectMany(static slide => slide.Pictures));
            Assert.Equal(0, Assert.Single(powerPoint.Report.EditablePages).OmittedImageCount);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfInvisibleImagePlacementSuppressed" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }
    }

    [Fact]
    public void ClippedImagePlacementsNeverEmbedRawPixelsAcrossEditableAdapters() {
        byte[] source = CreateDocument()
            .Canvas(canvas => canvas.Clip(
                40D,
                40D,
                20D,
                20D,
                clipped => clipped.Image(Png, 20D, 20D, 80D, 60D)))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);
        Assert.NotNull(placement.Clip);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Empty(word.Value.Images);
            Assert.True(word.HasLoss);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfImageClipNotSafelyEditable" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.DoesNotContain("data:image/", html.Value, StringComparison.Ordinal);
        Assert.True(html.HasLoss);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "ImageClipNotSafelyEditable" &&
            warning.LossKind == OfficeConversionLossKind.Omission);

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Empty(powerPoint.Value.Slides.SelectMany(static slide => slide.Pictures));
            Assert.True(powerPoint.HasLoss);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfImageClipNotSafelyEditable" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Fact]
    public void UnsupportedImageBlendModesNeverEmbedRawPixelsAcrossEditableAdapters() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImageWithGraphicsStatePdf("/BM /DefinitelyUnsupported"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);
        Assert.True(placement.HasUnsupportedBlendMode);
        Assert.False(placement.HasUnsupportedPaintState);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Empty(word.Value.Images);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfImageUnsupportedBlendModeNotSafelyEditable" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.DoesNotContain("data:image/", html.Value, StringComparison.Ordinal);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "ImageUnsupportedBlendModeNotSafelyEditable" &&
            warning.LossKind == OfficeConversionLossKind.Omission);

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Empty(powerPoint.Value.Slides.SelectMany(static slide => slide.Pictures));
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfImageUnsupportedBlendModeNotSafelyEditable" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Fact]
    public void NonBlendGraphicsStateEntriesDoNotMasqueradeAsUnsupportedBlendModes() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImageWithGraphicsStatePdf("/SA true /SM 0.02 /AIS false"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.False(placement.HasUnsupportedBlendMode);
        Assert.True(placement.HasUnsupportedPaintState);
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Single(word.Value.Images);
            Assert.DoesNotContain(word.Report.Warnings, static warning =>
                warning.Code == "PdfImageUnsupportedBlendModeNotSafelyEditable");
        }
    }

    [Fact]
    public void NonDefaultImageOpacityIsMappedAcrossEditableAdapters() {
        byte[] source = CreateDocument()
            .Canvas(canvas => canvas.Effect(
                OfficeIMO.Drawing.OfficeTransform.Identity,
                0.5D,
                effect => effect.Image(Png, 20D, 30D, 80D, 40D)))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Equal(50, Assert.Single(word.Value.Images).Transparency);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfImageOpacityMapped" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.Contains("style=\"opacity:0.5;\"", html.Value, StringComparison.Ordinal);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "ImageOpacityMapped" &&
            warning.LossKind == OfficeConversionLossKind.None);

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Equal(50, Assert.Single(Assert.Single(powerPoint.Value.Slides).Pictures).FillTransparency);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfImageOpacityMapped" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }
    }

    private static PdfDocument CreateDocument() => PdfDocument.Create(new PdfOptions {
        PageWidth = 160D,
        PageHeight = 160D,
        MarginLeft = 0D,
        MarginRight = 0D,
        MarginTop = 0D,
        MarginBottom = 0D
    });

    private static byte[] CreateImageWithGraphicsStatePdf(string graphicsStateEntries) {
        const string content = "q /GS1 gs 80 0 0 40 20 30 cm /Im1 Do Q\n";
        byte[] contentBytes = System.Text.Encoding.ASCII.GetBytes(content);
        byte[] imageBytes = { 255, 0, 0 };
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 160 160] /Resources << /XObject << /Im1 5 0 R >> /ExtGState << /GS1 6 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "endstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\n");
        output.Write(imageBytes, 0, imageBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Type /ExtGState " + graphicsStateEntries + " >>\nendobj\n");
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Size 7 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteAscii(Stream output, string value) {
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(value);
        output.Write(bytes, 0, bytes.Length);
    }
}

using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.Pdf.Word;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Producers such as LibreOffice draw every image inside a rectangular clip. A clip that contains the image keeps it
/// whole, a clip that cuts into it crops it to the visible pixels, and only clips that cannot be represented are
/// refused, always through the conversion report and never as text in the document body.
/// </summary>
public sealed class PdfClippedImageImportTests {
    private const int ImageSize = 10;
    private static readonly OfficeColor Red = OfficeColor.FromRgb(255, 0, 0);
    private static readonly OfficeColor Blue = OfficeColor.FromRgb(0, 0, 255);

    [Theory]
    [InlineData("20 30 80 40")]          // clip equals the image bounds
    [InlineData("0 0 160 160")]          // clip contains the image
    [InlineData("20.05 30.05 79.9 39.9")] // LibreOffice-style sub-pixel trim of the image frame
    public void ClipThatShowsTheWholeImageKeepsTheSourcePayloadInEveryTarget(string clip) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateSplitImagePdf(clip, "80 0 0 40 20 30"));
        byte[] source = Assert.Single(Assert.Single(logical.Pages).Images).SourceImage.Bytes;

        PdfWordConversionResult word = logical.ConvertToWordWithEvidence();
        using (word.Value) {
            WordImage image = Assert.Single(word.Value.Images);
            Assert.Equal(source, image.ToBytes());
            Assert.DoesNotContain(word.Report.Warnings, static warning => warning.Code == "pdf-image-clip-not-safely-editable");
        }

        AssertEveryTargetCarriesOneImage(logical, ImageSize, ImageSize, Red, Blue);
    }

    [Fact]
    public void ClipThatCutsIntoTheImageCropsItToTheVisiblePixels() {
        // The clip keeps the left half of the placement, which shows the red source columns.
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateSplitImagePdf("20 30 40 40", "80 0 0 40 20 30"));

        AssertEveryTargetCarriesOneImage(logical, ImageSize / 2, ImageSize, Red, Red);
    }

    [Fact]
    public void MarkdownAndHtmlExportEachDistinctCropOfOneImageResourceOnce() {
        // The same image is drawn three times: its left half, its right half, and its left half again elsewhere.
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImagePdf(
            "q 20 30 40 40 re W n q 80 0 0 40 20 30 cm /Im1 Do Q Q\n" +
            "q 60 90 40 40 re W n q 80 0 0 40 20 90 cm /Im1 Do Q Q\n" +
            "q 40 0 40 20 re W n q 80 0 0 20 40 0 cm /Im1 Do Q Q\n"));

        string html = logical.ToHtmlResult(PdfHtmlExportOptions.CreateSemanticProfile()).Value;
        foreach (string output in new[] { logical.ExtractMarkdown(), html }) {
            List<byte[]> images = ExtractDataUris(output).ToList();
            Assert.Equal(2, images.Count);
            Assert.Contains(images, static bytes => HasPixels(bytes, ImageSize / 2, ImageSize, Red, Red));
            Assert.Contains(images, static bytes => HasPixels(bytes, ImageSize / 2, ImageSize, Blue, Blue));
        }
    }

    [Fact]
    public void MarkdownImageFilesFromSeparateExportsShareADirectoryWithoutReplacingEachOther() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-markdown-" + Guid.NewGuid().ToString("N"));
        try {
            var options = new PdfLogicalMarkdownOptions { ImageExportMode = PdfLogicalMarkdownImageExportMode.File, ImageDirectory = directory };
            // Both documents name their only image the same way by page and position; only the pixels differ.
            string left = PdfDocumentReadResult.Load(CreateSplitImagePdf("20 30 40 40", "80 0 0 40 20 30")).ExtractMarkdown(options);
            string right = PdfDocumentReadResult.Load(CreateSplitImagePdf("60 30 40 40", "80 0 0 40 20 30")).ExtractMarkdown(options);

            AssertPixels(File.ReadAllBytes(Path.Combine(directory, GetImageFileName(left))), ImageSize / 2, ImageSize, Red, Red);
            AssertPixels(File.ReadAllBytes(Path.Combine(directory, GetImageFileName(right))), ImageSize / 2, ImageSize, Blue, Blue);
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void RtfPictureSizeAppliesThePageUserUnit() {
        // /UserUnit 2 makes the 80 x 40 unit placement (and its 40 x 40 unit crop) twice as large in physical points.
        foreach ((string clip, int widthTwips) in new[] { ("0 0 160 160", 80 * 2 * 20), ("20 30 40 40", 40 * 2 * 20) }) {
            PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImagePdf(
                "q " + clip + " re W n q 80 0 0 40 20 30 cm /Im1 Do Q Q\n", pageEntries: "/UserUnit 2"));

            using RtfDocument rtf = logical.ToRtf();
            RtfImage picture = Assert.IsType<RtfImage>(Assert.Single(rtf.Paragraphs.SelectMany(static paragraph => paragraph.Inlines)));
            Assert.Equal(widthTwips, picture.DesiredWidthTwips);
            Assert.Equal(40 * 2 * 20, picture.DesiredHeightTwips);
        }
    }

    [Fact]
    public void CropOfAMirroredImageKeepsTheSourcePixelsTheClipShows() {
        // A negative horizontal scale draws the blue source columns on the left of the placement.
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateSplitImagePdf("20 30 40 40", "-80 0 0 40 100 30"));

        PdfWordConversionResult word = logical.ConvertToWordWithEvidence();
        using (word.Value) {
            WordImage image = Assert.Single(word.Value.Images);
            Assert.True(image.HorizontalFlip);
            AssertPixels(image.ToBytes(), ImageSize / 2, ImageSize, Blue, Blue);
        }
    }

    [Theory]
    [InlineData(null, "20 30 m 100 30 l 60 70 l h W n")] // not a rectangle
    [InlineData("47.2 30 5.6 40", null)]                   // leaves 0.7 of a source pixel; each trimmed edge rounds down
    public void ClipThatCannotBeRepresentedIsReportedWithoutBodyText(string? clip, string? clipPath) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateSplitImagePdf(clip, "80 0 0 40 20 30", clipPath));

        PdfWordConversionResult word = logical.ConvertToWordWithEvidence();
        using (word.Value) {
            Assert.Empty(word.Value.Images);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "pdf-image-clip-not-safely-editable" && warning.LossKind == OfficeConversionLossKind.Omission);
            Assert.All(word.Value.Paragraphs, static paragraph => Assert.DoesNotContain("PDF image", paragraph.Text, StringComparison.Ordinal));
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfHtmlExportOptions.CreateSemanticProfile());
        Assert.DoesNotContain("<img", html.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("Im1", html.Value, StringComparison.Ordinal);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == OfficeDiagnosticCodes.ImageClipNotSafelyEditable && warning.LossKind == OfficeConversionLossKind.Omission);

        string markdown = logical.ExtractMarkdown();
        Assert.DoesNotContain("Image", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("Im1", markdown, StringComparison.Ordinal);

        using RtfDocument rtf = logical.ToRtf(null, out OfficeConversionReport rtfReport);
        Assert.DoesNotContain(rtf.Paragraphs.SelectMany(static paragraph => paragraph.Inlines), static inline => inline is RtfImage);
        Assert.DoesNotContain("Im1", rtf.ToRtf(), StringComparison.Ordinal);
        Assert.Contains(rtfReport.Diagnostics, static diagnostic =>
            diagnostic.Code == OfficeDiagnosticCodes.ImageClipNotSafelyEditable && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    private static void AssertEveryTargetCarriesOneImage(
        PdfDocumentReadResult logical,
        int width,
        int height,
        OfficeColor left,
        OfficeColor right) {
        PdfWordConversionResult word = logical.ConvertToWordWithEvidence();
        using (word.Value) {
            AssertPixels(Assert.Single(word.Value.Images).ToBytes(), width, height, left, right);
            Assert.All(word.Value.Paragraphs, static paragraph => Assert.DoesNotContain("PDF image", paragraph.Text, StringComparison.Ordinal));
        }

        using (OdtDocument odt = logical.ToOdt()) {
            OdtImage image = Assert.Single(odt.Paragraphs.SelectMany(static paragraph => paragraph.Images));
            AssertPixels(image.GetImageBytes(), width, height, left, right);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfHtmlExportOptions.CreateSemanticProfile());
        AssertPixels(Assert.Single(ExtractDataUris(html.Value)), width, height, left, right);
        Assert.DoesNotContain("<figcaption", html.Value, StringComparison.Ordinal);

        string markdown = logical.ExtractMarkdown();
        Assert.StartsWith("![](data:image/png;base64,", markdown, StringComparison.Ordinal);
        AssertPixels(Assert.Single(ExtractDataUris(markdown)), width, height, left, right);

        using (RtfDocument rtf = logical.ToRtf()) {
            RtfImage picture = Assert.IsType<RtfImage>(Assert.Single(rtf.Paragraphs.SelectMany(static paragraph => paragraph.Inlines)));
            Assert.Equal(RtfImageFormat.Png, picture.Format);
            AssertPixels(picture.Data, width, height, left, right);
        }

        PdfPowerPointConversionResult powerPoint = logical.ConvertToPowerPointWithEvidence(PdfPowerPointExportOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Single(powerPoint.Value.Slides.SelectMany(static slide => slide.Images));
        }
    }

    private static void AssertPixels(byte[] bytes, int width, int height, OfficeColor left, OfficeColor right) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? raster));
        Assert.NotNull(raster);
        Assert.Equal(width, raster!.Width);
        Assert.Equal(height, raster.Height);
        Assert.Equal(left, raster.GetPixel(0, height / 2));
        Assert.Equal(right, raster.GetPixel(width - 1, height / 2));
    }

    private static string GetImageFileName(string markdown) {
        int start = markdown.IndexOf("](", StringComparison.Ordinal) + 2;
        return markdown.Substring(start, markdown.IndexOf(')', start) - start);
    }

    private static bool HasPixels(byte[] bytes, int width, int height, OfficeColor left, OfficeColor right) =>
        OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? raster) &&
        raster!.Width == width && raster.Height == height &&
        raster.GetPixel(0, height / 2) == left && raster.GetPixel(width - 1, height / 2) == right;

    private static IEnumerable<byte[]> ExtractDataUris(string text) {
        const string Marker = ";base64,";
        int index = 0;
        while ((index = text.IndexOf("data:image/", index, StringComparison.Ordinal)) >= 0) {
            int start = text.IndexOf(Marker, index, StringComparison.Ordinal) + Marker.Length;
            int end = start;
            while (end < text.Length && (char.IsLetterOrDigit(text[end]) || text[end] is '+' or '/' or '=')) end++;
            yield return Convert.FromBase64String(text.Substring(start, end - start));
            index = end;
        }
    }

    /// <summary>A 160 pt page with one 10x10 RGB image: red left columns, blue right columns.</summary>
    private static byte[] CreateSplitImagePdf(string? clip, string matrix, string? clipPath = null) =>
        CreateImagePdf("q " + (clipPath ?? (clip == null ? string.Empty : clip + " re W n ")) + " q " + matrix + " cm /Im1 Do Q Q\n");

    /// <summary>A 160 pt page whose content stream draws the split image resource <c>/Im1</c>.</summary>
    private static byte[] CreateImagePdf(string content, string pageEntries = "") {
        byte[] contentBytes = System.Text.Encoding.ASCII.GetBytes(content);
        byte[] pixels = new byte[ImageSize * ImageSize * 3];
        for (int y = 0; y < ImageSize; y++) {
            for (int x = 0; x < ImageSize; x++) {
                int offset = (y * ImageSize + x) * 3;
                OfficeColor color = x < ImageSize / 2 ? Red : Blue;
                pixels[offset] = color.R;
                pixels[offset + 1] = color.G;
                pixels[offset + 2] = color.B;
            }
        }

        using var output = new MemoryStream();
        Write(output, "%PDF-1.7\n");
        Write(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        Write(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        Write(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 160 160] " + pageEntries + " /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        Write(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        Write(output, "endstream\nendobj\n");
        Write(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width " + ImageSize + " /Height " + ImageSize +
            " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length " + pixels.Length + " >>\nstream\n");
        output.Write(pixels, 0, pixels.Length);
        Write(output, "\nendstream\nendobj\n");
        Write(output, "trailer\n<< /Root 1 0 R /Size 6 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void Write(Stream stream, string text) {
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }
}

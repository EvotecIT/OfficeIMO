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

    [Theory]
    [InlineData("20 30 40 40", ImageSize / 2)] // the left half of the placement: the red source columns
    [InlineData("20 30 8 40", 1)]              // exactly one source pixel wide
    public void ClipThatCutsIntoTheImageCropsItToTheVisiblePixels(string clip, int width) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateSplitImagePdf(clip, "80 0 0 40 20 30"));

        AssertEveryTargetCarriesOneImage(logical, width, ImageSize, Red, Red);
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
    public void TextExtractionOrdersClippedImagesWithoutDecodingThem() {
        // The image's samples exceed the decoded-stream limit, so any attempt to decode them throws.
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxDecodedStreamBytes = 4096 } };
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateOversizedImagePdf(), null, options);

        Assert.Contains("Visible text", logical.Text, StringComparison.Ordinal);
        PdfLogicalImage image = Assert.Single(Assert.Single(logical.Pages).Images);
        Assert.Throws<PdfReadLimitException>(() => image.SourceImage.Bytes);
    }

    [Fact]
    public void HtmlEmbedsACropOfAJpeg2000ImageAsPng() {
        // Browsers cannot show JPEG 2000, but the crop is re-encoded as PNG and can be embedded.
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImagePdf(
            "q 20 30 40 40 re W n q 80 0 0 40 20 30 cm /Im1 Do Q Q\n", jpeg2000: true));

        string html = logical.ToHtmlResult(PdfHtmlExportOptions.CreateSemanticProfile()).Value;

        Assert.Contains("data:image/png;base64,", html, StringComparison.Ordinal);
        AssertPixels(Assert.Single(ExtractDataUris(html)), ImageSize / 2, ImageSize, Red, Red);
    }

    [Theory]
    [InlineData("0 0 160 160", ImageSize)]     // whole image: the JPEG is carried over
    [InlineData("20 30 40 40", ImageSize / 2)] // cropped: the visible pixels are re-encoded
    public void JpegExifOrientationDoesNotTurnTheExportedImage(string clip, int width) {
        // EXIF orientation 3 would turn the image 180 degrees and put the blue columns on the left.
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImagePdf(
            "q " + clip + " re W n q 80 0 0 40 20 30 cm /Im1 Do Q Q\n", jpegExifOrientation: 3));

        byte[] exported = Assert.Single(ExtractDataUris(logical.ExtractMarkdown()));

        // Decoded the way a browser shows it, with any EXIF orientation applied.
        Assert.True(OfficeRasterImageDecoder.TryDecode(exported, out OfficeRasterImage? raster));
        Assert.Equal(width, raster!.Width);
        Assert.Equal(ImageSize, raster.Height);
        OfficeColor pixel = raster.GetPixel(0, raster.Height / 2);
        Assert.True(pixel.R > 200 && pixel.B < 60, "Expected the red source columns on the left, got " + pixel + ".");
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

        string html = logical.ToHtmlResult(PdfHtmlExportOptions.CreateSemanticProfile()).Value;
        Assert.Contains("height=\"10\" style=\"transform:scale(-1,1);\">", html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("80 0 0 40 20 20", ImageSize)] // the lower placement sits wholly below the text
    [InlineData("80 0 0 120 20 20", 3)]        // it reaches above the text, but the clip keeps only its lowest rows
    public void EachCropKeepsItsReadingPositionAroundText(string lowerMatrix, int lowerHeight) {
        // The left half is drawn above the text and the right half below it.
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImagePdf(
            "q 20 110 40 40 re W n q 80 0 0 40 20 110 cm /Im1 Do Q Q\n" +
            "BT /F1 12 Tf 20 85 Td (Between the crops) Tj ET\n" +
            "q 60 20 40 40 re W n q " + lowerMatrix + " cm /Im1 Do Q Q\n"));

        string html = logical.ToHtmlResult(PdfHtmlExportOptions.CreateSemanticProfile()).Value;
        foreach (string output in new[] { logical.ExtractMarkdown(), html }) {
            int text = output.IndexOf("Between the crops", StringComparison.Ordinal);
            Assert.True(text > 0, output);
            Assert.True(HasPixels(Assert.Single(ExtractDataUris(output.Substring(0, text))), ImageSize / 2, ImageSize, Red, Red));
            Assert.True(HasPixels(Assert.Single(ExtractDataUris(output.Substring(text))), ImageSize / 2, lowerHeight, Blue, Blue));
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

    /// <summary>
    /// A 160 pt page whose content stream draws the split image resource <c>/Im1</c>, stored as raw RGB samples or,
    /// with <paramref name="jpegExifOrientation"/>, as a DCT image whose EXIF orientation PDF viewers ignore, or with
    /// <paramref name="jpeg2000"/>, as the same picture in the independent JPEG 2000 fixture.
    /// </summary>
    private static byte[] CreateImagePdf(string content, string pageEntries = "", int? jpegExifOrientation = null, bool jpeg2000 = false) {
        byte[] contentBytes = System.Text.Encoding.ASCII.GetBytes(content);
        var raster = new OfficeRasterImage(ImageSize, ImageSize, Blue);
        byte[] pixels = new byte[ImageSize * ImageSize * 3];
        for (int y = 0; y < ImageSize; y++) {
            for (int x = 0; x < ImageSize; x++) {
                int offset = (y * ImageSize + x) * 3;
                OfficeColor color = x < ImageSize / 2 ? Red : Blue;
                raster.SetPixel(x, y, color);
                pixels[offset] = color.R;
                pixels[offset + 1] = color.G;
                pixels[offset + 2] = color.B;
            }
        }
        string filter = string.Empty;
        if (jpegExifOrientation is { } orientation) {
            byte[] exif = {
                (byte)'I', (byte)'I', 42, 0, 8, 0, 0, 0,
                1, 0, 0x12, 1, 3, 0, 1, 0, 0, 0, (byte)orientation, 0, 0, 0, 0, 0, 0, 0
            };
            pixels = OfficeJpegCodec.Encode(raster, new OfficeJpegEncodeOptions {
                Quality = 95, Subsampling = OfficeJpegSubsampling.Y444, Metadata = new OfficeJpegMetadata(exif: exif)
            });
            filter = " /Filter /DCTDecode";
        } else if (jpeg2000) {
            pixels = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "Scans", "split-rgb.jp2"));
            filter = " /Filter /JPXDecode";
        }

        using var output = new MemoryStream();
        Write(output, "%PDF-1.7\n");
        Write(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        Write(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        Write(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 160 160] " + pageEntries + " /Resources << /XObject << /Im1 5 0 R >> /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> /Contents 4 0 R >>\nendobj\n");
        Write(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        Write(output, "endstream\nendobj\n");
        Write(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width " + ImageSize + " /Height " + ImageSize +
            " /ColorSpace /DeviceRGB /BitsPerComponent 8" + filter + " /Length " + pixels.Length + " >>\nstream\n");
        output.Write(pixels, 0, pixels.Length);
        Write(output, "\nendstream\nendobj\n");
        Write(output, "trailer\n<< /Root 1 0 R /Size 6 >>\n%%EOF\n");
        return output.ToArray();
    }

    /// <summary>A page with text and a clipped 100x100 RGB image whose zlib stream decodes to 30,000 bytes.</summary>
    private static byte[] CreateOversizedImagePdf() {
        byte[] content = System.Text.Encoding.ASCII.GetBytes(
            "q 20 30 40 40 re W n q 80 0 0 40 20 30 cm /Im1 Do Q Q\nBT /F1 12 Tf 20 120 Td (Visible text) Tj ET\n");
        byte[] samples = new byte[100 * 100 * 3];
        using var compressed = new MemoryStream();
        compressed.WriteByte(0x78);
        compressed.WriteByte(0x9C);
        using (var deflate = new System.IO.Compression.DeflateStream(compressed, System.IO.Compression.CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(samples, 0, samples.Length);
        }
        // Adler-32 of all-zero samples: the low sum stays 1 and the high sum is the sample count.
        int high = samples.Length % 65521;
        compressed.Write(new byte[] { (byte)(high >> 8), (byte)high, 0, 1 }, 0, 4);
        byte[] image = compressed.ToArray();

        using var output = new MemoryStream();
        Write(output, "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        Write(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        Write(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 160 160] /Resources << /XObject << /Im1 5 0 R >> " +
            "/Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> /Contents 4 0 R >>\nendobj\n");
        Write(output, "4 0 obj\n<< /Length " + content.Length + " >>\nstream\n");
        output.Write(content, 0, content.Length);
        Write(output, "endstream\nendobj\n");
        Write(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 100 /Height 100 /ColorSpace /DeviceRGB /BitsPerComponent 8 " +
            "/Filter /FlateDecode /Length " + image.Length + " >>\nstream\n");
        output.Write(image, 0, image.Length);
        Write(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R /Size 6 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void Write(Stream stream, string text) {
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }
}

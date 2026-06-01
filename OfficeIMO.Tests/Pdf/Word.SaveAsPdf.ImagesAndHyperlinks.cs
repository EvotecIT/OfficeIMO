using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf_ImagesAndHyperlinks() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfImagesLinks.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfImagesLinks.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph().AddImage(imagePath, 50, 50);
            document.AddHyperLink("OfficeIMO", new Uri("https://evotec.xyz"), addStyle: true);
            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/URI (https://evotec.xyz", pdfContent);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Aligned_Images() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAlignedImages.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAlignedImages.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordParagraph left = document.AddParagraph();
            left.ParagraphAlignment = JustificationValues.Left;
            left.AddImage(imagePath, 48, 48);

            WordParagraph center = document.AddParagraph();
            center.ParagraphAlignment = JustificationValues.Center;
            center.AddImage(imagePath, 48, 48);

            WordParagraph right = document.AddParagraph();
            right.ParagraphAlignment = JustificationValues.Right;
            right.AddImage(imagePath, 48, 48);

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new OfficeIMO.Pdf.PageSize(300, 260),
                Margins = OfficeIMO.Pdf.PageMargins.Uniform(30)
            });
        }

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Match mediaBox = Regex.Match(pdfContent, @"/MediaBox\s*\[0 0 (?<width>\d+(?:\.\d+)?) (?<height>\d+(?:\.\d+)?)\]");
        Assert.True(mediaBox.Success, "Expected generated PDF to expose a simple MediaBox.");

        double pageWidth = double.Parse(mediaBox.Groups["width"].Value, CultureInfo.InvariantCulture);
        const double margin = 30D;
        const double imageWidth = 36D;
        double[] imageXPositions = Regex.Matches(pdfContent, @"36 0 0 36 (?<x>-?\d+(?:\.\d+)?) -?\d+(?:\.\d+)? cm\s*/Im\d+ Do")
            .Cast<Match>()
            .Select(match => double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture))
            .ToArray();

        Assert.True(imageXPositions.Length >= 3, "Expected three native image placement matrices.");
        Assert.InRange(imageXPositions[0], margin - 1D, margin + 1D);
        Assert.InRange(imageXPositions[1], margin + ((pageWidth - (2D * margin) - imageWidth) / 2D) - 1D, margin + ((pageWidth - (2D * margin) - imageWidth) / 2D) + 1D);
        Assert.InRange(imageXPositions[2], pageWidth - margin - imageWidth - 1D, pageWidth - margin - imageWidth + 1D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Exports_Loaded_Inline_Paragraph_Images() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeLoadedInlineImage.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeLoadedInlineImage.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordParagraph paragraph = document.AddParagraph();
            paragraph.AddText("Before loaded image ");
            paragraph.AddImage(imagePath, 48, 48);
            paragraph.AddText(" after loaded image");
            document.Save();
        }

        using (WordDocument document = WordDocument.Load(docPath)) {
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyImageUnsupported");
        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("/Subtype /Image", pdfContent);

        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Before loaded image", text);
        Assert.Contains("after loaded image", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Skips_Loaded_Unsupported_Png_Images() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeLoadedUnsupportedPng.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeLoadedUnsupportedPng.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Before unsupported image");
            document.AddParagraph().AddImage(imagePath, 48, 48);
            document.AddParagraph("After unsupported image");
            document.Save();
        }

        ReplaceFirstMainDocumentImagePart(docPath, CreateUnsupportedInterlacedPng());

        using (WordDocument document = WordDocument.Load(docPath)) {
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.Contains(options.Warnings, warning =>
            warning.Code == "NativeBodyImageUnsupported" &&
            warning.Message.Contains("PNG", StringComparison.OrdinalIgnoreCase));
        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Before unsupported image", text);
        Assert.Contains("After unsupported image", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Body_PictureControl_To_Image() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativePictureControl.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativePictureControl.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Logo content control:");
            WordParagraph picture = document.AddParagraph();
            picture.ParagraphAlignment = JustificationValues.Center;
            picture.AddPictureControl(imagePath, 48, 48, "Logo", "LogoTag");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body paragraph");

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("36 0 0 36", pdfContent);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Cell_PictureControl_To_Image() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellPictureControl.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellPictureControl.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 7200;
            table.ColumnWidth = new[] { 7200 }.ToList();
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "Logo content control in a table";
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.AddPictureControl(imagePath, 48, 48, "Cell Logo", "CellLogo");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body table");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        string pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("36 0 0 36", pdfContent);

        using PdfDocument pdf = PdfDocument.Open(bytes);
        Assert.Contains("Logo content control in a table", pdf.GetPage(1).Text);
    }

    private static void ReplaceFirstMainDocumentImagePart(string docPath, byte[] bytes) {
        using WordprocessingDocument package = WordprocessingDocument.Open(docPath, true);
        ImagePart imagePart = package.MainDocumentPart!.ImageParts.First();
        using var stream = new MemoryStream(bytes);
        imagePart.FeedData(stream);
    }

    private static byte[] CreateUnsupportedInterlacedPng() {
        byte[] bytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAABGdBTAAAACklEQVR42mP8z8AABQMBgA4uA1sAAAAASUVORK5CYII=");
        bytes[28] = 1;
        return bytes;
    }
}

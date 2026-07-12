using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Worksheet_Images() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfImages.xlsx");

        byte[] imageBytes = CreateMinimalRgbPng();
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "ImageMarker");
            sheet.AddImage(2, 1, imageBytes, "image/png", widthPixels: 24, heightPixels: 16, name: "Pdf Logo", altText: "PDF logo");

            ExcelImage image = Assert.Single(sheet.Images);
            Assert.Equal("Pdf Logo", image.Name);
            Assert.Equal("PDF logo", image.Description);
            Assert.Equal(2, image.RowIndex);
            Assert.Equal(1, image.ColumnIndex);
            Assert.Equal(24, image.WidthPixels);
            Assert.Equal(16, image.HeightPixels);
            Assert.Equal("image/png", image.ContentType);
            Assert.Equal(imageBytes, image.GetBytes());

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("ImageMarker", pdf.GetPage(1).Text);

        var extractedImages = PdfCore.PdfImageExtractor.ExtractImages(bytes);
        var extractedImage = Assert.Single(extractedImages);
        Assert.Equal(1, extractedImage.PageNumber);
        Assert.Equal("png", extractedImage.FileExtension);
        Assert.Equal("image/png", extractedImage.MimeType);
        Assert.True(extractedImage.IsImageFile);
        Assert.Equal(1, extractedImage.Width);
        Assert.Equal(1, extractedImage.Height);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Preserves_Worksheet_Image_Rotation() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfRotatedImage.xlsx");

        byte[] imageBytes = CreateMinimalRgbPng();
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "ImageMarker");
            sheet.AddImage(2, 1, imageBytes, "image/png", widthPixels: 24, heightPixels: 16, name: "Rotated Logo").SetRotation(30);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0
            });
        }

        PdfCore.PdfImagePlacement placement = Assert.Single(PdfCore.PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.False(placement.IsAxisAligned);
        Assert.True(placement.B < 0);
        Assert.True(placement.C > 0);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Filters_Images_Anchored_To_Hidden_Cells() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHiddenCellImages.xlsx");

        byte[] imageBytes = CreateMinimalRgbPng();
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "VisibleMarker");
            sheet.Cell(2, 1, "HiddenImageMarker");
            sheet.AddImage(2, 1, imageBytes, "image/png", widthPixels: 24, heightPixels: 16, name: "Hidden Logo", altText: "Hidden logo");
            sheet.SetRowHidden(2, true);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                RespectWorksheetHiddenRowsAndColumns = true
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("VisibleMarker", text);
        Assert.DoesNotContain("HiddenImageMarker", text);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(bytes));
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Warns_And_Skips_Invalid_Worksheet_Image_Bytes() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfInvalidImageBytes.xlsx");
        byte[] invalidPngBytes = new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            16, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "ImageMarker");
            sheet.AddImage(2, 1, invalidPngBytes, "image/png", widthPixels: 24, heightPixels: 16, name: "Invalid PNG");
            document.Save();

            bytes = document.ToPdf(options);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("ImageMarker", pdf.GetPage(1).Text);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(bytes));
        Assert.Contains(options.Warnings, warning => warning.SheetName == "Images" && warning.Feature == "WorksheetImage");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Warns_And_Skips_Worksheet_Image_When_Declared_Type_Differs_From_Bytes() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfMismatchedImageType.xlsx");
        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "ImageMarker");
            sheet.AddImage(2, 1, CreateMinimalRgbPng(), "image/jpeg", widthPixels: 24, heightPixels: 16, name: "Declared JPEG");
            document.Save();

            bytes = document.ToPdf(options);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("ImageMarker", pdf.GetPage(1).Text);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(bytes));
        Assert.Contains(options.Warnings, warning =>
            warning.SheetName == "Images" &&
            warning.Feature == "WorksheetImage" &&
            warning.Message.Contains("Image bytes were declared as JPEG but were detected as Png.", StringComparison.Ordinal));
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Embeds_Worksheet_Images_In_Anchored_Table_Cells() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfAnchoredImageCell.xlsx");

        byte[] imageBytes = CreateMinimalRgbPng();
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Label");
            sheet.Cell(1, 2, "Visual");
            sheet.Cell(2, 1, "BeforeImageRow");
            sheet.Cell(3, 1, "AnchoredImageRow");
            sheet.Cell(4, 1, "AfterImageRow");
            sheet.AddImage(3, 2, imageBytes, "image/png", widthPixels: 72, heightPixels: 72, name: "Anchored Cell Image");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(420, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        UglyToad.PdfPig.Content.Page page = pdf.GetPage(1);
        double beforeRowY = FindWordStartY(page, "BeforeImageRow");
        double anchoredRowY = FindWordStartY(page, "AnchoredImageRow");
        double afterRowY = FindWordStartY(page, "AfterImageRow");

        double gapBeforeAnchoredRow = beforeRowY - anchoredRowY;
        double gapAfterAnchoredRow = anchoredRowY - afterRowY;
        Assert.True(gapAfterAnchoredRow > gapBeforeAnchoredRow + 20, "The image should increase the anchored table row height instead of flowing before the table.");

        var extractedImage = Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(bytes));
        Assert.Equal(1, extractedImage.PageNumber);
        Assert.Equal("image/png", extractedImage.MimeType);
    }

}

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
            Assert.Equal(imageBytes, image.ToBytes());

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
        PdfCore.PdfDocumentConversionResult result;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "ImageMarker");
            sheet.AddImage(2, 1, invalidPngBytes, "image/png", widthPixels: 24, heightPixels: 16, name: "Invalid PNG");
            document.Save();

            result = document.ToPdfDocumentResult(options);
            bytes = result.ToBytes();
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("ImageMarker", pdf.GetPage(1).Text);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(bytes));
        Assert.Contains(result.Warnings, warning => warning.Source == "Images" && warning.Code == "WorksheetImage");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Warns_And_Skips_Worksheet_Image_When_Declared_Type_Differs_From_Bytes() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfMismatchedImageType.xlsx");
        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false
        };

        byte[] bytes;
        PdfCore.PdfDocumentConversionResult result;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "ImageMarker");
            sheet.AddImage(2, 1, CreateMinimalRgbPng(), "image/jpeg", widthPixels: 24, heightPixels: 16, name: "Declared JPEG");
            document.Save();

            result = document.ToPdfDocumentResult(options);
            bytes = result.ToBytes();
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("ImageMarker", pdf.GetPage(1).Text);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(bytes));
        Assert.Contains(result.Warnings, warning =>
            warning.Source == "Images" &&
            warning.Code == "WorksheetImage" &&
            warning.Message.Contains("Image bytes were declared as JPEG but were detected as Png.", StringComparison.Ordinal));
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Positions_Worksheet_Images_At_Drawing_Anchors() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfAnchoredImageCell.xlsx");

        byte[] imageBytes = CreateMinimalRgbPng();
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Label");
            sheet.Cell(1, 2, "Visual");
            sheet.Cell(2, 1, "BeforeRow");
            sheet.Cell(3, 1, "AnchorRow");
            sheet.Cell(4, 1, "AfterRow");
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
        double beforeRowY = FindWordStartY(page, "BeforeRow");
        double anchoredRowY = FindWordStartY(page, "AnchorRow");
        double afterRowY = FindWordStartY(page, "AfterRow");

        double gapBeforeAnchoredRow = beforeRowY - anchoredRowY;
        double gapAfterAnchoredRow = anchoredRowY - afterRowY;
        Assert.InRange(Math.Abs(gapAfterAnchoredRow - gapBeforeAnchoredRow), 0D, 2D);

        var extractedImage = Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(bytes));
        Assert.Equal(1, extractedImage.PageNumber);
        Assert.Equal("image/png", extractedImage.MimeType);
        PdfCore.PdfImagePlacement placement = Assert.Single(PdfCore.PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.InRange(placement.X, 70D, 74D);
        Assert.InRange(placement.Y, 250D, 254D);
        Assert.InRange(placement.Width, 53D, 55D);
        Assert.InRange(placement.Height, 53D, 55D);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_FlowTable_Mode_Embeds_Images_In_Anchored_Cells() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfFlowAnchoredImageCell.xlsx");
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Label");
            sheet.Cell(1, 2, "Visual");
            sheet.Cell(2, 1, "BeforeImageRow");
            sheet.Cell(3, 1, "AnchoredImageRow");
            sheet.Cell(4, 1, "AfterImageRow");
            sheet.AddImage(3, 2, CreateMinimalRgbPng(), "image/png", widthPixels: 72, heightPixels: 72, name: "Anchored Cell Image");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable,
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

        Assert.True(
            anchoredRowY - afterRowY > beforeRowY - anchoredRowY + 20D,
            "Flow compatibility mode should retain the historic in-cell image layout.");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Renders_A_Paginated_Image_Only_On_Its_Owning_Page() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfPaginatedImage.xlsx");
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            for (int row = 1; row <= 45; row++) {
                sheet.Cell(row, 1, "Row " + row.ToString(CultureInfo.InvariantCulture));
            }
            sheet.AddImage(42, 2, CreateMinimalRgbPng(), "image/png", widthPixels: 48, heightPixels: 32, name: "Late page image");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(300, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1);
        PdfCore.PdfExtractedImage image = Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(bytes));
        Assert.Equal(pdf.NumberOfPages, image.PageNumber);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Positions_Terminal_Media_After_The_Final_Body_Chunk() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfTerminalMedia.xlsx");
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            for (int row = 1; row <= 45; row++) {
                sheet.Cell(row, 1, "Row " + row.ToString(CultureInfo.InvariantCulture));
            }
            sheet.AddImage(55, 2, CreateMinimalRgbPng(), "image/png", widthPixels: 72, heightPixels: 32, name: "Terminal image");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(300, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        PdfCore.PdfImagePlacement image = Assert.Single(PdfCore.PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(pdf.NumberOfPages, image.PageNumber);
        Assert.True(
            image.Width > 20D,
            "Terminal media should be scaled against the final body chunk, not the entire repeated-header worksheet.");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Preserves_Separate_Anchors_On_A_Media_Only_Sheet() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfMediaOnlyAnchors.xlsx");
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            byte[] image = CreateMinimalRgbPng();
            sheet.AddImage(2, 2, image, "image/png", widthPixels: 24, heightPixels: 16, name: "Upper image");
            sheet.AddImage(8, 4, image, "image/png", widthPixels: 24, heightPixels: 16, name: "Lower image");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 300),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        IReadOnlyList<PdfCore.PdfImagePlacement> placements = PdfCore.PdfImageExtractor.ExtractImagePlacements(bytes);
        Assert.Equal(2, placements.Count);
        Assert.True(Math.Abs(placements[0].X - placements[1].X) > 20D);
        Assert.True(Math.Abs(placements[0].Y - placements[1].Y) > 20D);
    }

}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData("LH", false, true)]
    [InlineData("LH", true, true)]
    [InlineData("RF", false, true)]
    [InlineData("LH", true, false)]
    public void SaveAsPdf_HeaderFooterVml_Resolves_Image_Relationships(
        string shapeId, bool declareRelationshipNamespace, bool staleRelationshipId) {
        byte[] logo = CreateMinimalRgbPng();
        using var workbook = new MemoryStream();
        using (var package = SpreadsheetDocument.Create(workbook, SpreadsheetDocumentType.Workbook, true)) {
            WorkbookPart workbookPart = package.AddWorkbookPart();
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("sheet");
            workbookPart.Workbook = new Workbook(new Sheets(new Sheet {
                Id = "sheet", SheetId = 1U, Name = "Logo"
            }));
            var headerFooter = new HeaderFooter();
            if (shapeId == "LH") headerFooter.OddHeader = new OddHeader("&L&GDaily Excel");
            else headerFooter.OddFooter = new OddFooter("&R&G");
            worksheetPart.Worksheet = new Worksheet(new SheetData(new Row(new Cell {
                CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("Logo body")
            }) { RowIndex = 1U }), headerFooter, new LegacyDrawingHeaderFooter { Id = "vml" });

            VmlDrawingPart vmlPart = worksheetPart.AddNewPart<VmlDrawingPart>("vml");
            ImagePart imagePart = vmlPart.AddImagePart(ImagePartType.Png, "rId1");
            using (var imageStream = new MemoryStream(logo)) imagePart.FeedData(imageStream);

            // x1's VML has an undeclared r prefix and a stale r:id alongside a valid o:relid.
            // Build the imported representation directly; a writer round-trip cannot expose this defect.
            string namespaceDeclaration = declareRelationshipNamespace
                ? "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"" : string.Empty;
            string imageReference = staleRelationshipId
                ? "r:id=\"R6c94117169f741dc\" o:relid=\"rId1\"" : "r:id=\"rId1\"";
            string vml = $"<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" {namespaceDeclaration}>" +
                $"<v:shape id=\"{shapeId}\" o:spid=\"_x0000_s1025\" style=\"width:28pt;height:14pt\">" +
                $"<v:imagedata {imageReference}/></v:shape></xml>";
            using var vmlStream = new MemoryStream(Encoding.UTF8.GetBytes(vml));
            vmlPart.FeedData(vmlStream);
        }

        workbook.Position = 0;
        using ExcelDocument document = ExcelDocument.Load(workbook, new ExcelLoadOptions {
            AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly
        });
        ExcelSheet.ExcelHeaderFooterInfo snapshot = document.Sheets[0].GetHeaderFooter();
        ExcelSheet.ExcelHeaderFooterImageInfo? image = shapeId == "LH" ? snapshot.HeaderLeftImage : snapshot.FooterRightImage;
        Assert.NotNull(image);
        Assert.Equal(logo, image.Bytes);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(28D, image.Width.ToPoints());
        Assert.Equal(14D, image.Height.ToPoints());
        Assert.Equal(shapeId == "LH" ? ExcelHeaderFooterPosition.Left : ExcelHeaderFooterPosition.Right, image.Position);

        byte[] pdf = document.ToPdf(new ExcelPdfExportOptions { IncludeSheetHeadings = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.Equal(1, Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(pdf)).PageNumber);
    }
}

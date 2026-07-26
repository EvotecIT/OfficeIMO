using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void ToPdfDocument_ProjectsBuiltInExcelTableStyleIntoPdfCells() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfStructuredTableStyle.xlsx");

        PdfCore.PdfDocument pdfDocument;
        PdfCore.PdfColor expectedAccent;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Report")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(10, 8, "Column1");
            sheet.Cell(10, 9, "SamAccountName");
            sheet.Cell(10, 10, "Group Name");
            sheet.Cell(11, 8, "DE-IPH-SQLDIST1");
            sheet.Cell(11, 9, "gMSADE189Acz5SS$");
            sheet.Cell(11, 10, "gMSA-de-DE-IPH-SQLDIST1_TR_Distribution");
            sheet.AddTable(
                "H10:J11",
                hasHeader: true,
                name: "IdentityData",
                style: TableStyle.TableStyleLight9,
                includeAutoFilter: true);
            document.Save();
            string accentArgb = Assert.IsType<string>(document.ResolveThemeColorArgb(4U));
            expectedAccent = PdfCore.PdfColor.FromRgb(
                Convert.ToByte(accentArgb.Substring(accentArgb.Length - 6, 2), 16),
                Convert.ToByte(accentArgb.Substring(accentArgb.Length - 4, 2), 16),
                Convert.ToByte(accentArgb.Substring(accentArgb.Length - 2, 2), 16));

            pdfDocument = document.ToPdfDocument(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable
            });
        }

        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(pdfDocument.Blocks));
        PdfCore.TableBlock table = Assert.Single(page.Blocks.OfType<PdfCore.TableBlock>());
        PdfCore.PdfTableStyle style = Assert.IsType<PdfCore.PdfTableStyle>(table.Style);

        Assert.NotNull(style.CellFills);
        Assert.Equal(expectedAccent, style.CellFills![(0, 0)]);
        Assert.Equal(expectedAccent, style.CellFills[(0, 2)]);
        Assert.False(style.CellFills.ContainsKey((1, 0)));
        Assert.NotNull(style.CellBorders);
        Assert.Equal(6, style.CellBorders!.Count);

        PdfCore.TextRun headerRun = Assert.Single(table.Cells[0][0].Runs);
        Assert.True(headerRun.Bold);
        Assert.Equal(PdfCore.PdfColor.FromRgb(255, 255, 255), headerRun.Color);
        Assert.False(Assert.Single(table.Cells[1][0].Runs).Bold);
    }

    [Fact]
    public void ToPdfDocument_BoundedReadPreservesDirectStylesWithoutLoadingWorksheetDom() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfBoundedStreaming.xlsx");
        using (ExcelDocument source = ExcelDocument.Create(workbookPath, "Report")) {
            ExcelSheet sheet = source.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(1, 2, "Status");
            sheet.Cell(2, 1, "Platform");
            sheet.Cell(2, 2, "Ready");
            sheet.Cell(3, 1, "Ignored");
            sheet.Cell(3, 2, "Outside limit");
            sheet.Range("A1").SetFillColor("C00000").SetFontColor("FFFFFF").SetBold();
            source.Save();
        }

        using ExcelDocument document = ExcelDocument.Load(
            workbookPath,
            new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        var worksheetPart = Assert.Single(document.OpenXmlDocument.WorkbookPart!.WorksheetParts);
        Assert.False(worksheetPart.IsRootElementLoaded);

        PdfCore.PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            MaxRowsPerSheet = 2,
            UseBoundedWorksheetRead = true,
            WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable
        });

        Assert.False(worksheetPart.IsRootElementLoaded);
        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "WorksheetBoundedRead");
        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(conversion.Value.Blocks));
        PdfCore.TableBlock table = Assert.Single(page.Blocks.OfType<PdfCore.TableBlock>());
        PdfCore.PdfTableStyle style = Assert.IsType<PdfCore.PdfTableStyle>(table.Style);
        Assert.Equal(PdfCore.PdfColor.FromRgb(192, 0, 0), style.CellFills![(0, 0)]);
        PdfCore.TextRun headerRun = Assert.Single(table.Cells[0][0].Runs);
        Assert.True(headerRun.Bold);
        Assert.Equal(PdfCore.PdfColor.FromRgb(255, 255, 255), headerRun.Color);
        Assert.Equal(2, table.Rows.Count);
    }

    [Fact]
    public void ToPdfDocument_BoundedReadFallsBackToCreatedWorksheetStyles() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1, "Name");
        sheet.Cell(2, 1, "Platform");
        sheet.Range("A1").SetFillColor("C00000").SetFontColor("FFFFFF").SetBold();

        PdfCore.PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            MaxRowsPerSheet = 2,
            UseBoundedWorksheetRead = true,
            WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable
        });

        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(conversion.Value.Blocks));
        PdfCore.TableBlock table = Assert.Single(page.Blocks.OfType<PdfCore.TableBlock>());
        PdfCore.PdfTableStyle style = Assert.IsType<PdfCore.PdfTableStyle>(table.Style);
        Assert.Equal(PdfCore.PdfColor.FromRgb(192, 0, 0), style.CellFills![(0, 0)]);
        PdfCore.TextRun headerRun = Assert.Single(table.Cells[0][0].Runs);
        Assert.True(headerRun.Bold);
        Assert.Equal(PdfCore.PdfColor.FromRgb(255, 255, 255), headerRun.Color);
    }

    [Fact]
    public void ToPdfDocument_BoundedReadIgnoresStaleDefaultWorksheetDimension() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfBoundedStaleDimension.xlsx");
        using (ExcelDocument source = ExcelDocument.Create(workbookPath, "Report")) {
            source.Sheets[0].Cell(3, 2, "Preserved");
            source.Save();
        }

        using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet =
               DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(workbookPath, true)) {
            DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet =
                spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
            worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetDimension>()!.Reference = "A1";
            worksheet.Save();
        }

        using ExcelDocument document = ExcelDocument.Load(
            workbookPath,
            new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        PdfCore.PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            MaxRowsPerSheet = 2,
            UseBoundedWorksheetRead = true,
            WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable
        });

        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(conversion.Value.Blocks));
        PdfCore.TableBlock table = Assert.Single(page.Blocks.OfType<PdfCore.TableBlock>());
        Assert.Equal("Preserved", Assert.Single(table.Cells[0][0].Runs).Text);
    }
}

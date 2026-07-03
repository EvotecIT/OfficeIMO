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
    public void SaveAsPdf_ExcelWorkbook_Uses_Worksheet_Column_Widths_And_Print_Scale() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfColumnWidths.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Widths")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "ARef");
            sheet.Cell(1, 2, "WideColumn");
            sheet.Cell(1, 3, "Tail");
            sheet.SetColumnWidth(1, 8);
            sheet.SetColumnWidth(2, 32);
            sheet.SetColumnWidth(3, 8);
            sheet.SetPageSetup(scale: 50);

            IReadOnlyList<ExcelColumnSnapshot> columns = sheet.GetColumnDefinitions();
            Assert.Equal(3, columns.Count);
            Assert.Equal(32, columns[1].Width);
            Assert.True(columns[1].CustomWidth);
            Assert.Equal((uint)50, sheet.GetPageSetup().Scale);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("ARef", text);
        Assert.Contains("WideColumn", text);
        Assert.Contains("Tail", text);

        double firstColumnX = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Min(letter => letter.StartBaseLine.X);
        double wideColumnX = FindFirstLetterStartX(page, "W");
        double tailX = FindFirstLetterStartX(page, "T");
        Assert.True(tailX - wideColumnX > (wideColumnX - firstColumnX) * 2D, $"Expected worksheet column width proportions to make the middle column visibly wider. A: {firstColumnX:0.##}, B: {wideColumnX:0.##}, C: {tailX:0.##}.");
        Assert.True(tailX < 190D, $"Expected worksheet print scale to narrow the rendered table. Tail x: {tailX:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Worksheet_Row_Heights() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfRowHeights.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Heights")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "TopRow");
            sheet.Cell(2, 1, "TallRow");
            sheet.Cell(3, 1, "AfterTall");
            sheet.SetRowHeight(2, 60);

            ExcelRowSnapshot row = Assert.Single(sheet.GetRowDefinitions());
            Assert.Equal(2, row.Index);
            Assert.Equal(60, row.Height);
            Assert.True(row.CustomHeight);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(260, 260),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("TopRow", text);
        Assert.Contains("TallRow", text);
        Assert.Contains("AfterTall", text);

        double topY = FindWordStartY(page, "TopRow");
        double tallY = FindWordStartY(page, "TallRow");
        double afterY = FindWordStartY(page, "AfterTall");
        double defaultGap = topY - tallY;
        double customGap = tallY - afterY;
        Assert.True(customGap > defaultGap * 2D, $"Expected worksheet row height to create a visibly taller second PDF table row. Default gap: {defaultGap:0.##}, custom gap: {customGap:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Omits_Hidden_Rows_And_Columns() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHiddenRowsColumns.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "VisibleOnly")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "VisibleHeader");
            sheet.Cell(1, 2, "HiddenColumnValue");
            sheet.Cell(2, 1, "HiddenRowValue");
            sheet.Cell(3, 1, "VisibleTail");
            sheet.SetColumnHidden(2, true);
            sheet.SetRowHidden(2, true);

            ExcelColumnSnapshot column = Assert.Single(sheet.GetColumnDefinitions());
            Assert.Equal(2, column.StartIndex);
            Assert.Equal(2, column.EndIndex);
            Assert.True(column.Hidden);

            ExcelRowSnapshot row = Assert.Single(sheet.GetRowDefinitions());
            Assert.Equal(2, row.Index);
            Assert.True(row.Hidden);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(320, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("VisibleHeader", text);
        Assert.Contains("VisibleTail", text);
        Assert.DoesNotContain("HiddenColumnValue", text);
        Assert.DoesNotContain("HiddenRowValue", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Merged_Cells_To_Table_Spans() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfMergedCells.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Merged")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "MergedTitle");
            sheet.Cell(1, 3, "TailCell");
            sheet.Cell(2, 1, "ColumnA");
            sheet.Cell(2, 2, "ColumnB");
            sheet.Cell(2, 3, "ColumnC");
            sheet.MergeRange("A1:B1");

            ExcelMergedRangeSnapshot mergedRange = Assert.Single(sheet.GetMergedRanges());
            Assert.Equal("A1:B1", mergedRange.A1Range);
            Assert.Equal(1, mergedRange.StartRow);
            Assert.Equal(1, mergedRange.StartColumn);
            Assert.Equal(1, mergedRange.EndRow);
            Assert.Equal(2, mergedRange.EndColumn);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("MergedTitle", text);
        Assert.Contains("TailCell", text);
        Assert.Contains("ColumnA", text);
        Assert.Contains("ColumnB", text);
        Assert.Contains("ColumnC", text);

        double mergedTitleX = FindWordStartX(page, "MergedTitle");
        double tailCellX = FindWordStartX(page, "TailCell");
        double columnBX = FindWordStartX(page, "ColumnB");
        double columnCX = FindWordStartX(page, "ColumnC");

        Assert.True(tailCellX > columnBX + 30D, $"Expected tail cell after A1:B1 merge to render in the third visual column. Tail x: {tailCellX:0.##}, ColumnB x: {columnBX:0.##}.");
        Assert.InRange(tailCellX, columnCX - 4D, columnCX + 4D);
        Assert.True(mergedTitleX < columnBX, $"Expected merged title to start in the first visual column. Title x: {mergedTitleX:0.##}, ColumnB x: {columnBX:0.##}.");
    }

}

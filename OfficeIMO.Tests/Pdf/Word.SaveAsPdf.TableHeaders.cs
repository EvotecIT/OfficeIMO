using DocumentFormat.OpenXml.Wordprocessing;
using System;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Header_Row() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRepeatingHeaderTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRepeatingHeaderTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(46, 2);
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    cell.Width = 1440;
                    cell.WidthType = TableWidthUnitValues.Dxa;
                }
            }

            table.Rows[0].Cells[0].Paragraphs[0].Text = "RepeatHdr";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "ValueHdr";
            table.RepeatAsHeaderRowAtTheTopOfEachPage = true;

            for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++) {
                table.Rows[rowIndex].Cells[0].Paragraphs[0].Text = "Row " + rowIndex.ToString("D2");
                table.Rows[rowIndex].Cells[1].Paragraphs[0].Text = "Value " + rowIndex.ToString("D2");
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(260, 220),
                Margins = PdfCore.PageMargins.Uniform(12)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);

        int repeatedHeaderCount = pdf.GetPages()
            .SelectMany(page => page.GetWords())
            .Count(word => word.Text == "RepeatHdr");
        Assert.True(repeatedHeaderCount >= 2);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Multiple_Header_Rows() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeMultipleRepeatingHeaderRows.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeMultipleRepeatingHeaderRows.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(44, 2);
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    cell.Width = 1440;
                    cell.WidthType = TableWidthUnitValues.Dxa;
                }
            }

            table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
            table.Rows[1].RepeatHeaderRowAtTheTopOfEachPage = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "HdrA";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "HdrB";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "HdrC";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "HdrD";

            for (int rowIndex = 2; rowIndex < table.Rows.Count; rowIndex++) {
                table.Rows[rowIndex].Cells[0].Paragraphs[0].Text = "Metric " + rowIndex.ToString("D2", System.Globalization.CultureInfo.InvariantCulture);
                table.Rows[rowIndex].Cells[1].Paragraphs[0].Text = "Owner " + rowIndex.ToString("D2", System.Globalization.CultureInfo.InvariantCulture);
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(260, 220),
                Margins = PdfCore.PageMargins.Uniform(12)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);

        int firstHeaderCount = pdf.GetPages()
            .SelectMany(page => page.GetWords())
            .Count(word => word.Text == "HdrA");
        int secondHeaderCount = pdf.GetPages()
            .SelectMany(page => page.GetWords())
            .Count(word => word.Text == "HdrC");

        Assert.True(firstHeaderCount >= 2);
        Assert.True(secondHeaderCount >= 2);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_First_Row_Style_Without_Repeating() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeFirstRowStyleNoRepeat.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeFirstRowStyleNoRepeat.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(44, 2, WordTableStyle.GridTable1Light);
            table.ConditionalFormattingFirstRow = true;
            table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = false;
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    cell.Width = 1440;
                    cell.WidthType = TableWidthUnitValues.Dxa;
                }
            }

            table.Rows[0].Cells[0].Paragraphs[0].Text = "SoloHdr";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "SoloValue";
            for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++) {
                table.Rows[rowIndex].Cells[0].Paragraphs[0].Text = "Body " + rowIndex.ToString("D2", System.Globalization.CultureInfo.InvariantCulture);
                table.Rows[rowIndex].Cells[1].Paragraphs[0].Text = "Value " + rowIndex.ToString("D2", System.Globalization.CultureInfo.InvariantCulture);
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(260, 220),
                Margins = PdfCore.PageMargins.Uniform(12)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);

        int firstRowHeaderCount = pdf.GetPages()
            .SelectMany(page => page.GetWords())
            .Count(word => word.Text == "SoloHdr");
        Assert.Equal(1, firstRowHeaderCount);
    }

}

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
    public void Test_WordDocument_SaveAsPdf_TableStyles() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfStyledTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfStyledTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Styled";
            cell.ShadingFillColorHex = "FF0000";
            cell.Borders.TopStyle = BorderValues.Single;
            cell.Borders.BottomStyle = BorderValues.Single;
            cell.Borders.LeftStyle = BorderValues.Single;
            cell.Borders.RightStyle = BorderValues.Single;
            cell.Borders.TopColorHex = "0000FF";
            cell.Borders.BottomColorHex = "0000FF";
            cell.Borders.LeftColorHex = "0000FF";
            cell.Borders.RightColorHex = "0000FF";
            cell.Borders.TopSize = 8;
            cell.Borders.BottomSize = 8;
            cell.Borders.LeftSize = 8;
            cell.Borders.RightSize = 8;
            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        string content = ReadPdfPageContent(bytes);
        Assert.Contains("1 0 0 rg", content);
        Assert.Contains("0 0 1 RG", content);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_Fill_And_Border() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeStyledTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeStyledTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Native styled cell";
            cell.ShadingFillColorHex = "ff0000";
            cell.Borders.TopStyle = BorderValues.Single;
            cell.Borders.BottomStyle = BorderValues.Single;
            cell.Borders.LeftStyle = BorderValues.Single;
            cell.Borders.RightStyle = BorderValues.Single;
            cell.Borders.TopColorHex = "0000ff";
            cell.Borders.BottomColorHex = "0000ff";
            cell.Borders.LeftColorHex = "0000ff";
            cell.Borders.RightColorHex = "0000ff";
            cell.Borders.TopSize = 8;
            cell.Borders.BottomSize = 8;
            cell.Borders.LeftSize = 8;
            cell.Borders.RightSize = 8;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Native styled cell", text);
        }

        string content = ReadPdfPageContent(bytes);
        Assert.Contains("1 0 0 rg", content);
        Assert.Contains("0 0 1 RG", content);
        Assert.Contains("1 w", content);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_NonUniform_Borders() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellNonUniformBorders.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellNonUniformBorders.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Native non-uniform cell";
            cell.Borders.TopStyle = BorderValues.Single;
            cell.Borders.TopColorHex = "ff0000";
            cell.Borders.TopSize = 16;
            cell.Borders.RightStyle = BorderValues.Single;
            cell.Borders.RightColorHex = "0000ff";
            cell.Borders.RightSize = 20;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new OfficeIMO.Pdf.PageSize(360, 200),
                Margins = OfficeIMO.Pdf.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Native non-uniform", text);
            Assert.Contains("cell", text);
        }

        string content = ReadPdfPageContent(bytes);
        Assert.Contains("1 0 0 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("0 0 1 RG", content);
        Assert.Contains("2.5 w", content);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_Double_And_Diagonal_Borders() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellDoubleDiagonalBorders.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellDoubleDiagonalBorders.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Native double diagonal cell";
            cell.Borders.TopStyle = BorderValues.Double;
            cell.Borders.BottomStyle = BorderValues.Double;
            cell.Borders.LeftStyle = BorderValues.Double;
            cell.Borders.RightStyle = BorderValues.Double;
            cell.Borders.TopColorHex = "123456";
            cell.Borders.BottomColorHex = "123456";
            cell.Borders.LeftColorHex = "123456";
            cell.Borders.RightColorHex = "123456";
            cell.Borders.TopSize = 8;
            cell.Borders.BottomSize = 8;
            cell.Borders.LeftSize = 8;
            cell.Borders.RightSize = 8;
            cell.Borders.TopLeftToBottomRightStyle = BorderValues.Double;
            cell.Borders.TopLeftToBottomRightColorHex = "654321";
            cell.Borders.TopLeftToBottomRightSize = 8;
            cell.Borders.TopRightToBottomLeftStyle = BorderValues.Double;
            cell.Borders.TopRightToBottomLeftColorHex = "654321";
            cell.Borders.TopRightToBottomLeftSize = 8;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new OfficeIMO.Pdf.PageSize(360, 200),
                Margins = OfficeIMO.Pdf.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Native double diagonal", text);
        }

        string content = ReadPdfPageContent(bytes);
        Assert.Contains("0.071 0.204 0.337 RG", content, StringComparison.Ordinal);
        Assert.Contains("0.396 0.263 0.129 RG", content, StringComparison.Ordinal);
        Assert.True(content.Split(new[] { " S" }, StringSplitOptions.None).Length - 1 >= 10, "Expected Word double and diagonal borders to emit multiple stroked lines.");
        Assert.True(content.Contains(" m ", StringComparison.Ordinal) && content.Contains(" l S", StringComparison.Ordinal), "Expected Word diagonal borders to emit PDF line segments.");
    }
}

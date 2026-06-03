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
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Word_Style() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordStyledTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordStyledTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(2, 2);
            table.Style = WordTableStyle.TableGrid;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Styled grid";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Header value";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Body label";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Body value";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Styled grid", text);
            Assert.Contains("Body value", text);
        }

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0 0 0 RG", raw);
        Assert.Contains("0.5 w", raw);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Level_Borders() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableLevelBorders.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableLevelBorders.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(2, 2);
            table.Style = WordTableStyle.PlainTable1;
            table.StyleDetails!.SetBordersForAllSides(BorderValues.Single, 12U, OfficeIMO.Drawing.OfficeColor.Red);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Border A1";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Border B1";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Border A2";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Border B2";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Border A1", text);
            Assert.Contains("Border B2", text);
        }

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("1 0 0 RG", raw);
        Assert.Contains("1.5 w", raw);
    }
}

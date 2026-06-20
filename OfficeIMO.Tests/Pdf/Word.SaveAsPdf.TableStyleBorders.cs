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

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Custom_Table_Style_Borders() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCustomTableStyleBorders.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCustomTableStyleBorders.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeCustomTableStyleBorders";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Custom Table Style Borders" },
                new StyleTableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Color = "008000", Size = 12U },
                        new LeftBorder { Val = BorderValues.Single, Color = "008000", Size = 12U },
                        new BottomBorder { Val = BorderValues.Single, Color = "008000", Size = 12U },
                        new RightBorder { Val = BorderValues.Single, Color = "008000", Size = 12U },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Color = "008000", Size = 12U },
                        new InsideVerticalBorder { Val = BorderValues.Single, Color = "008000", Size = 12U })))
            {
                Type = StyleValues.Table,
                StyleId = styleId,
                CustomStyle = true
            });

            WordTable table = document.AddTable(2, 2);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Style border A1";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Style border B2";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Style border A1", text);
            Assert.Contains("Style border B2", text);
        }

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0 0.502 0 RG", raw);
        Assert.Contains("1.5 w", raw);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Custom_Table_Style_NonUniform_Borders() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCustomTableStyleNonUniformBorders.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCustomTableStyleNonUniformBorders.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeCustomTableStyleNonUniformBorders";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Custom Table Style Non Uniform Borders" },
                new StyleTableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Color = "FF0000", Size = 16U },
                        new LeftBorder { Val = BorderValues.Single, Color = "000000", Size = 8U },
                        new BottomBorder { Val = BorderValues.Single, Color = "0000FF", Size = 20U },
                        new RightBorder { Val = BorderValues.Single, Color = "000000", Size = 8U },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Color = "008000", Size = 8U },
                        new InsideVerticalBorder { Val = BorderValues.Single, Color = "FFFF00", Size = 12U })))
            {
                Type = StyleValues.Table,
                StyleId = styleId,
                CustomStyle = true
            });

            WordTable table = document.AddTable(2, 2);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Style nonuniform A1";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Style nonuniform B1";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Style nonuniform A2";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Style nonuniform B2";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Style nonuniform A1", text);
            Assert.Contains("Style nonuniform B2", text);
        }

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("1 0 0 RG", raw);
        Assert.Contains("2 w", raw);
        Assert.Contains("0 0 1 RG", raw);
        Assert.Contains("2.5 w", raw);
        Assert.Contains("0 0.502 0 RG", raw);
        Assert.Contains("1 1 0 RG", raw);
    }
}

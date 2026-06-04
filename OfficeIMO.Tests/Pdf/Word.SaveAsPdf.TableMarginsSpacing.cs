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
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Default_Cell_Margins() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellMargins.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellMargins.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable plain = document.AddTable(1, 1);
            ConfigureMarginTable(plain, "PlainPad");
            plain.StyleDetails!.MarginDefaultLeftWidth = 0;

            document.AddParagraph("between padding tables");

            WordTable padded = document.AddTable(1, 1);
            ConfigureMarginTable(padded, "WidePad");
            padded.StyleDetails!.MarginDefaultLeftWidth = 1000;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(400, 500),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var plainWord = Assert.Single(words, word => word.Text == "PlainPad");
        var paddedWord = Assert.Single(words, word => word.Text == "WidePad");

        Assert.True(paddedWord.BoundingBox.Left > plainWord.BoundingBox.Left + 35D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_Margins() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTablePerCellMargins.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTablePerCellMargins.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable plain = document.AddTable(1, 1);
            ConfigureMarginTable(plain, "PlainCellPad");
            plain.StyleDetails!.MarginDefaultLeftWidth = 0;

            document.AddParagraph("between per-cell padding tables");

            WordTable padded = document.AddTable(1, 1);
            ConfigureMarginTable(padded, "WideCellPad");
            padded.StyleDetails!.MarginDefaultLeftWidth = 0;
            padded.Rows[0].Cells[0].MarginLeftWidth = 1000;
            padded.Rows[0].Cells[0].MarginTopWidth = 320;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(400, 500),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var plainWord = Assert.Single(words, word => word.Text == "PlainCellPad");
        var paddedWord = Assert.Single(words, word => word.Text == "WideCellPad");

        Assert.True(paddedWord.BoundingBox.Left > plainWord.BoundingBox.Left + 35D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_Spacing() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellSpacing.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellSpacing.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable plain = document.AddTable(1, 2);
            ConfigureCellSpacingTable(plain, "PlainA", "PlainB");
            plain.StyleDetails!.CellSpacing = 0;

            document.AddParagraph("between spacing tables");

            WordTable spaced = document.AddTable(1, 2);
            ConfigureCellSpacingTable(spaced, "SpaceA", "SpaceB");
            spaced.StyleDetails!.CellSpacing = 240;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(420, 500),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var plainLeft = Assert.Single(words, word => word.Text == "PlainA");
        var plainRight = Assert.Single(words, word => word.Text == "PlainB");
        var spacedLeft = Assert.Single(words, word => word.Text == "SpaceA");
        var spacedRight = Assert.Single(words, word => word.Text == "SpaceB");

        double plainGap = plainRight.BoundingBox.Left - plainLeft.BoundingBox.Left;
        double spacedGap = spacedRight.BoundingBox.Left - spacedLeft.BoundingBox.Left;
        Assert.True(spacedGap > plainGap + 10D, $"Expected Word table cell spacing to widen native table cell distance. Plain gap: {plainGap}; spaced gap: {spacedGap}.");
    }
}

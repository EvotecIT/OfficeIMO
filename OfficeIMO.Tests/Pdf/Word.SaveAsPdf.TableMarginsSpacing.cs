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

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Cell_Spacing() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleCellSpacing.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleCellSpacing.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Generic Rendered Spaced Table" },
                new StyleTableProperties(new TableCellSpacing {
                    Width = "240",
                    Type = TableWidthUnitValues.Dxa
                }))
            { Type = StyleValues.Table, StyleId = "GenericRenderedSpacedTable" });

            WordTable plain = document.AddTable(1, 2);
            ConfigureCellSpacingTable(plain, "PA", "PB");
            plain.StyleDetails!.CellSpacing = 0;

            document.AddParagraph("between style spacing tables");

            WordTable spaced = document.AddTable(1, 2);
            ConfigureCellSpacingTable(spaced, "SA", "SB");
            spaced._tableProperties!.TableStyle = new TableStyle { Val = "GenericRenderedSpacedTable" };

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
        var plainLeft = Assert.Single(words, word => word.Text == "PA");
        var plainRight = Assert.Single(words, word => word.Text == "PB");
        var spacedLeft = Assert.Single(words, word => word.Text == "SA");
        var spacedRight = Assert.Single(words, word => word.Text == "SB");

        double plainGap = plainRight.BoundingBox.Left - plainLeft.BoundingBox.Left;
        double spacedGap = spacedRight.BoundingBox.Left - spacedLeft.BoundingBox.Left;
        Assert.True(spacedGap > plainGap + 10D,
            $"Expected Word table style cell spacing to widen native table cell distance. Plain gap: {plainGap:0.##}; spaced gap: {spacedGap:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Conditional_Cell_Margins() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleConditionalCellMargins.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleConditionalCellMargins.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Generic Rendered Conditional Margin Table" },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new TableCellMargin(
                            new LeftMargin { Width = "720", Type = TableWidthUnitValues.Dxa })))
                { Type = TableStyleOverrideValues.FirstColumn })
            { Type = StyleValues.Table, StyleId = "GenericRenderedConditionalMarginTable" });

            WordTable plain = document.AddTable(2, 3);
            ConfigureCellSpacingTable(plain, "PlainLeft", "PlainMiddle");
            plain.Rows[0].Cells[2].Paragraphs[0].Text = "PlainRight";

            document.AddParagraph("between conditional margin tables");

            WordTable padded = document.AddTable(2, 3);
            ConfigureCellSpacingTable(padded, "PaddedLeft", "PaddedMiddle");
            padded.Rows[0].Cells[2].Paragraphs[0].Text = "PaddedRight";
            padded._tableProperties!.TableStyle = new TableStyle { Val = "GenericRenderedConditionalMarginTable" };
            padded.ConditionalFormattingFirstColumn = true;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(520, 360),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var plainLeft = Assert.Single(words, word => word.Text == "PlainLeft");
        var paddedLeft = Assert.Single(words, word => word.Text == "PaddedLeft");

        Assert.True(paddedLeft.BoundingBox.Left > plainLeft.BoundingBox.Left + 25D,
            $"Expected first-column conditional cell margin to move text right. Plain x: {plainLeft.BoundingBox.Left:0.##}; padded x: {paddedLeft.BoundingBox.Left:0.##}.");
    }
}

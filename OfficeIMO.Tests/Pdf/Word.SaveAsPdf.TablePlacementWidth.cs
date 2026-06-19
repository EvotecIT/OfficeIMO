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
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Placement() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTablePlacement.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTablePlacement.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable leftTable = document.AddTable(1, 2);
            ConfigurePlacementTable(leftTable, "LeftTbl", TableRowAlignmentValues.Left);

            document.AddParagraph("between left and center");

            WordTable centerTable = document.AddTable(1, 2);
            ConfigurePlacementTable(centerTable, "CenterTbl", TableRowAlignmentValues.Center);

            document.AddParagraph("between center and right");

            WordTable rightTable = document.AddTable(1, 2);
            ConfigurePlacementTable(rightTable, "RightTbl", TableRowAlignmentValues.Right);

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
        var left = Assert.Single(words, word => word.Text == "LeftTbl");
        var center = Assert.Single(words, word => word.Text == "CenterTbl");
        var right = Assert.Single(words, word => word.Text == "RightTbl");

        Assert.True(center.BoundingBox.Left > left.BoundingBox.Left + 70D);
        Assert.True(right.BoundingBox.Left > center.BoundingBox.Left + 70D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Alignment() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleAlignment.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleAlignment.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Generic Centered Table" },
                new StyleTableProperties(new TableJustification {
                    Val = TableRowAlignmentValues.Center
                }))
            { Type = StyleValues.Table, StyleId = "GenericCenteredTable" });

            WordTable leftTable = document.AddTable(1, 2);
            ConfigurePlacementTable(leftTable, "StyleLeft", TableRowAlignmentValues.Left);

            document.AddParagraph("between styled alignment tables");

            WordTable styledTable = document.AddTable(1, 2);
            foreach (WordTableCell cell in styledTable.Rows[0].Cells) {
                cell.Width = 1440;
                cell.WidthType = TableWidthUnitValues.Dxa;
            }

            styledTable.Rows[0].Cells[0].Paragraphs[0].Text = "StyleCenter";
            styledTable.Rows[0].Cells[1].Paragraphs[0].Text = "Value";
            styledTable._tableProperties!.TableStyle = new TableStyle { Val = "GenericCenteredTable" };

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(400, 360),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var left = Assert.Single(words, word => word.Text == "StyleLeft");
        var center = Assert.Single(words, word => word.Text == "StyleCenter");

        Assert.True(center.BoundingBox.Left > left.BoundingBox.Left + 70D,
            $"Expected Word table style alignment to center the native table. Left x: {left.BoundingBox.Left:0.##}; centered x: {center.BoundingBox.Left:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Preferred_Dxa_Width() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTablePreferredWidth.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTablePreferredWidth.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable preferred = document.AddTable(1, 2);
            preferred.LayoutType = TableLayoutValues.Fixed;
            preferred.WidthType = TableWidthUnitValues.Dxa;
            preferred.Width = 2160;
            preferred.Rows[0].Cells[0].Paragraphs[0].Text = "NA";
            preferred.Rows[0].Cells[1].Paragraphs[0].Text = "NB";

            document.AddParagraph("between width tables");

            WordTable defaultWidth = document.AddTable(1, 2);
            defaultWidth.Rows[0].Cells[0].Paragraphs[0].Text = "FA";
            defaultWidth.Rows[0].Cells[1].Paragraphs[0].Text = "FB";

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
        var narrowLeft = Assert.Single(words, word => word.Text == "NA");
        var narrowRight = Assert.Single(words, word => word.Text == "NB");
        var defaultLeft = Assert.Single(words, word => word.Text == "FA");
        var defaultRight = Assert.Single(words, word => word.Text == "FB");

        double preferredGap = narrowRight.BoundingBox.Left - narrowLeft.BoundingBox.Left;
        double defaultGap = defaultRight.BoundingBox.Left - defaultLeft.BoundingBox.Left;
        Assert.True(preferredGap < defaultGap - 40D, $"Expected preferred DXA table width to narrow the native table. Preferred gap: {preferredGap}; default gap: {defaultGap}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Left_Indent() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleRenderedLeftIndent.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleRenderedLeftIndent.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Generic Rendered Indented Table" },
                new StyleTableProperties(new TableIndentation {
                    Width = 720,
                    Type = TableWidthUnitValues.Dxa
                }))
            { Type = StyleValues.Table, StyleId = "GenericRenderedIndentedTable" });

            WordTable defaultTable = document.AddTable(1, 1);
            ConfigureMarginTable(defaultTable, "DefaultStyleIndent");

            document.AddParagraph("between style indent tables");

            WordTable styledTable = document.AddTable(1, 1);
            ConfigureMarginTable(styledTable, "StyledIndent");
            styledTable._tableProperties!.TableStyle = new TableStyle { Val = "GenericRenderedIndentedTable" };

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(400, 280),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var defaultWord = Assert.Single(words, word => word.Text == "DefaultStyleIndent");
        var styledWord = Assert.Single(words, word => word.Text == "StyledIndent");

        Assert.True(styledWord.BoundingBox.Left > defaultWord.BoundingBox.Left + 30D,
            $"Expected table style indentation to move the styled native table right. Default x: {defaultWord.BoundingBox.Left:0.##}; styled x: {styledWord.BoundingBox.Left:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Preferred_Width() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleRenderedPreferredWidth.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleRenderedPreferredWidth.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Generic Rendered Width Table" },
                new StyleTableProperties(new TableWidth {
                    Width = "1440",
                    Type = TableWidthUnitValues.Dxa
                }))
            { Type = StyleValues.Table, StyleId = "GenericRenderedWidthTable" });

            WordTable defaultTable = document.AddTable(1, 2);
            ConfigureCellSpacingTable(defaultTable, "DA", "DB");

            document.AddParagraph("between styled width tables");

            WordTable styledTable = document.AddTable(1, 2);
            ConfigureCellSpacingTable(styledTable, "SA", "SB");
            styledTable._tableProperties!.TableStyle = new TableStyle { Val = "GenericRenderedWidthTable" };

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(400, 280),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var defaultLeft = Assert.Single(words, word => word.Text == "DA");
        var defaultRight = Assert.Single(words, word => word.Text == "DB");
        var styledLeft = Assert.Single(words, word => word.Text == "SA");
        var styledRight = Assert.Single(words, word => word.Text == "SB");

        double defaultGap = defaultRight.BoundingBox.Left - defaultLeft.BoundingBox.Left;
        double styledGap = styledRight.BoundingBox.Left - styledLeft.BoundingBox.Left;
        Assert.True(styledGap < defaultGap - 25D,
            $"Expected table style preferred width to narrow the styled native table. Default gap: {defaultGap:0.##}; styled gap: {styledGap:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Autofit_Layout() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleAutofitLayout.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleAutofitLayout.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Generic Rendered Autofit Table" },
                new StyleTableProperties(new DocumentFormat.OpenXml.Wordprocessing.TableLayout {
                    Type = TableLayoutValues.Autofit
                }))
            { Type = StyleValues.Table, StyleId = "GenericRenderedAutofitTable" });

            WordTable fixedTable = document.AddTable(1, 2);
            fixedTable.LayoutType = TableLayoutValues.Fixed;
            fixedTable.Rows[0].Cells[0].Width = 2880;
            fixedTable.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Dxa;
            fixedTable.Rows[0].Cells[1].Width = 2880;
            fixedTable.Rows[0].Cells[1].WidthType = TableWidthUnitValues.Dxa;
            fixedTable.Rows[0].Cells[0].Paragraphs[0].Text = "FA";
            fixedTable.Rows[0].Cells[1].Paragraphs[0].Text = "FixedWide";

            document.AddParagraph("between style layout tables");

            WordTable styledTable = document.AddTable(1, 2);
            styledTable._tableProperties!.TableStyle = new TableStyle { Val = "GenericRenderedAutofitTable" };
            styledTable._tableProperties.TableLayout?.Remove();
            styledTable.Rows[0].Cells[0].Width = 2880;
            styledTable.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Dxa;
            styledTable.Rows[0].Cells[1].Width = 2880;
            styledTable.Rows[0].Cells[1].WidthType = TableWidthUnitValues.Dxa;
            styledTable.Rows[0].Cells[0].Paragraphs[0].Text = "SA";
            styledTable.Rows[0].Cells[1].Paragraphs[0].Text = "StyledWide";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(420, 300),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var fixedLeft = Assert.Single(words, word => word.Text == "FA");
        var fixedRight = Assert.Single(words, word => word.Text == "FixedWide");
        var styledLeft = Assert.Single(words, word => word.Text == "SA");
        var styledRight = Assert.Single(words, word => word.Text == "StyledWide");

        double fixedGap = fixedRight.BoundingBox.Left - fixedLeft.BoundingBox.Left;
        double styledGap = styledRight.BoundingBox.Left - styledLeft.BoundingBox.Left;
        Assert.True(styledGap < fixedGap - 25D,
            $"Expected table style autofit layout to shrink the first-column gap. Fixed gap: {fixedGap:0.##}; styled gap: {styledGap:0.##}.");
    }
}

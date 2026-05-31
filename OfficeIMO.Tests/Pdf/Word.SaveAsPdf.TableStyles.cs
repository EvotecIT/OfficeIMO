using DocumentFormat.OpenXml.Wordprocessing;
using System;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using UglyToad.PdfPig;
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
        string content = ReadFirstPdfStreamContent(bytes);
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
        using (PdfDocument pdf = PdfDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Native styled cell", text);
        }

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("1 0 0 rg", raw);
        Assert.Contains("0 0 1 RG", raw);
        Assert.Contains("1 w", raw);
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
                OfficeIMOPageSize = new OfficeIMO.Pdf.PageSize(360, 200),
                OfficeIMOMargins = OfficeIMO.Pdf.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfDocument pdf = PdfDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Native non-uniform", text);
            Assert.Contains("cell", text);
        }

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("1 0 0 RG", raw);
        Assert.Contains("2 w", raw);
        Assert.Contains("0 0 1 RG", raw);
        Assert.Contains("2.5 w", raw);
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
                OfficeIMOPageSize = new OfficeIMO.Pdf.PageSize(360, 200),
                OfficeIMOMargins = OfficeIMO.Pdf.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfDocument pdf = PdfDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Native double diagonal", text);
        }

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.071 0.204 0.337 RG", raw, StringComparison.Ordinal);
        Assert.Contains("0.396 0.263 0.129 RG", raw, StringComparison.Ordinal);
        Assert.True(raw.Split(new[] { " S" }, StringSplitOptions.None).Length - 1 >= 10, "Expected Word double and diagonal borders to emit multiple stroked lines.");
        Assert.True(raw.Contains(" m ", StringComparison.Ordinal) && raw.Contains(" l S", StringComparison.Ordinal), "Expected Word diagonal borders to emit PDF line segments.");
    }

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
        using (PdfDocument pdf = PdfDocument.Open(bytes)) {
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
        using (PdfDocument pdf = PdfDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Border A1", text);
            Assert.Contains("Border B2", text);
        }

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("1 0 0 RG", raw);
        Assert.Contains("1.5 w", raw);
    }

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
                OfficeIMOPageSize = new PdfCore.PageSize(400, 500),
                OfficeIMOMargins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
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
                OfficeIMOPageSize = new PdfCore.PageSize(400, 500),
                OfficeIMOMargins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
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
                OfficeIMOPageSize = new PdfCore.PageSize(420, 500),
                OfficeIMOMargins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
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
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Row_Height() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableRowHeight.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableRowHeight.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            table.Rows[0].Height = 1600;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "TallRow";
            document.AddParagraph("AfterTall");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                OfficeIMOPageSize = new PdfCore.PageSize(400, 500),
                OfficeIMOMargins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var tableWord = Assert.Single(words, word => word.Text == "TallRow");
        var followingWord = Assert.Single(words, word => word.Text == "AfterTall");

        Assert.True(tableWord.BoundingBox.Bottom > followingWord.BoundingBox.Top + 45D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_NonUniform_Row_Heights() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableNonUniformRowHeights.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableNonUniformRowHeights.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(3, 1);
            table.Rows[0].Height = 400;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "ShortA";
            table.Rows[1].Height = 1200;
            table.Rows[1].Cells[0].Paragraphs[0].Text = "TallB";
            table.Rows[2].Height = 400;
            table.Rows[2].Cells[0].Paragraphs[0].Text = "ShortC";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                OfficeIMOPageSize = new PdfCore.PageSize(320, 260),
                OfficeIMOMargins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var shortA = Assert.Single(words, word => word.Text == "ShortA");
        var tallB = Assert.Single(words, word => word.Text == "TallB");
        var shortC = Assert.Single(words, word => word.Text == "ShortC");

        double firstGap = shortA.BoundingBox.Bottom - tallB.BoundingBox.Bottom;
        double secondGap = tallB.BoundingBox.Bottom - shortC.BoundingBox.Bottom;
        Assert.True(secondGap > firstGap + 35D, $"Expected non-uniform Word row height to push the third row down. ShortA/TallB gap: {firstGap}; TallB/ShortC gap: {secondGap}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_Hyperlink() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeLinkedTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeLinkedTable.pdf");
        const string linkUri = "https://evotec.xyz/native-table-link";

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].AddHyperLink("Native table link", new Uri(linkUri), addStyle: true, tooltip: "Native table link metadata");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfDocument pdf = PdfDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Native table link", text);
        }

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);
        Assert.Equal(linkUri, link.Uri);
        Assert.Equal("Native table link metadata", link.Contents);
        Assert.Equal(new[] { linkUri }, info.LinkUris);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_Alignment() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAlignedTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAlignedTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 3);
            WordTableCell leftCell = table.Rows[0].Cells[0];
            WordTableCell centerCell = table.Rows[0].Cells[1];
            WordTableCell rightCell = table.Rows[0].Cells[2];

            leftCell.Width = 1440;
            leftCell.WidthType = TableWidthUnitValues.Dxa;
            centerCell.Width = 1440;
            centerCell.WidthType = TableWidthUnitValues.Dxa;
            rightCell.Width = 1440;
            rightCell.WidthType = TableWidthUnitValues.Dxa;

            leftCell.Paragraphs[0].Text = "TOP";
            leftCell.AddParagraph("PAD");
            leftCell.AddParagraph("PAD");
            leftCell.Paragraphs[0].ParagraphAlignment = JustificationValues.Left;
            leftCell.Paragraphs[1].ParagraphAlignment = JustificationValues.Left;
            leftCell.Paragraphs[2].ParagraphAlignment = JustificationValues.Left;
            leftCell.VerticalAlignment = TableVerticalAlignmentValues.Top;

            centerCell.Paragraphs[0].Text = "MID";
            centerCell.Paragraphs[0].ParagraphAlignment = JustificationValues.Center;
            centerCell.VerticalAlignment = TableVerticalAlignmentValues.Center;

            rightCell.Paragraphs[0].Text = "END";
            rightCell.Paragraphs[0].ParagraphAlignment = JustificationValues.Right;
            rightCell.VerticalAlignment = TableVerticalAlignmentValues.Bottom;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var top = Assert.Single(words, word => word.Text == "TOP");
        var mid = Assert.Single(words, word => word.Text == "MID");
        var end = Assert.Single(words, word => word.Text == "END");

        const double columnWidth = 72D;
        double firstColumnLeft = top.BoundingBox.Left - 4D;
        double secondColumnLeft = firstColumnLeft + columnWidth;
        double thirdColumnLeft = secondColumnLeft + columnWidth;

        Assert.InRange(top.BoundingBox.Left, firstColumnLeft + 3D, firstColumnLeft + 8D);
        Assert.InRange(mid.BoundingBox.Left, secondColumnLeft + 20D, secondColumnLeft + 36D);
        Assert.InRange(end.BoundingBox.Right, thirdColumnLeft + columnWidth - 8D, thirdColumnLeft + columnWidth - 2D);
        Assert.True(top.BoundingBox.Bottom > mid.BoundingBox.Bottom + 8D);
        Assert.True(mid.BoundingBox.Bottom > end.BoundingBox.Bottom + 8D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_NonUniform_Alignment() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeNonUniformAlignedTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeNonUniformAlignedTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(2, 2);
            foreach (WordTableRow row in table.Rows) {
                row.Height = 1100;
                foreach (WordTableCell cell in row.Cells) {
                    cell.Width = 1440;
                    cell.WidthType = TableWidthUnitValues.Dxa;
                }
            }

            table.Rows[0].Cells[0].Paragraphs[0].Text = "TopPeer";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Left2";
            table.Rows[0].Cells[1].Paragraphs[0].ParagraphAlignment = JustificationValues.Left;
            table.Rows[0].Cells[1].VerticalAlignment = TableVerticalAlignmentValues.Top;

            table.Rows[1].Cells[0].Paragraphs[0].Text = "TopCell";
            table.Rows[1].Cells[0].VerticalAlignment = TableVerticalAlignmentValues.Top;
            table.Rows[1].Cells[1].Paragraphs[0].Text = "R2";
            table.Rows[1].Cells[1].Paragraphs[0].ParagraphAlignment = JustificationValues.Right;
            table.Rows[1].Cells[1].VerticalAlignment = TableVerticalAlignmentValues.Bottom;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                OfficeIMOPageSize = new PdfCore.PageSize(360, 260),
                OfficeIMOMargins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var leftPeer = Assert.Single(words, word => word.Text == "Left2");
        var topCell = Assert.Single(words, word => word.Text == "TopCell");
        var rightBottom = Assert.Single(words, word => word.Text == "R2");

        Assert.True(rightBottom.BoundingBox.Left > leftPeer.BoundingBox.Left + 35D, $"Expected non-uniform right-aligned cell to move right. Left2 x: {leftPeer.BoundingBox.Left}; R2 x: {rightBottom.BoundingBox.Left}.");
        Assert.True(topCell.BoundingBox.Bottom > rightBottom.BoundingBox.Bottom + 20D, $"Expected non-uniform bottom-aligned cell to move down inside the same row. TopCell bottom: {topCell.BoundingBox.Bottom}; R2 bottom: {rightBottom.BoundingBox.Bottom}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Merged_Cells() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeMergedTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeMergedTable.pdf");
        const string horizontalUri = "https://evotec.xyz/native-table-column-span";
        const string verticalUri = "https://evotec.xyz/native-table-row-span";

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(3, 3);
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    cell.Width = 1440;
                    cell.WidthType = TableWidthUnitValues.Dxa;
                }
            }

            table.Rows[0].Cells[0].Paragraphs[0].AddHyperLink("Across", new Uri(horizontalUri), addStyle: true, tooltip: "Column span metadata");
            table.Rows[0].Cells[0].Paragraphs[0].ParagraphAlignment = JustificationValues.Center;
            table.Rows[0].Cells[2].Paragraphs[0].Text = "TopTail";

            table.Rows[1].Cells[0].Paragraphs[0].AddHyperLink("Tall", new Uri(verticalUri), addStyle: true, tooltip: "Row span metadata");
            table.Rows[1].Cells[0].VerticalAlignment = TableVerticalAlignmentValues.Center;
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Upper";
            table.Rows[1].Cells[2].Paragraphs[0].Text = "UpperTail";
            table.Rows[2].Cells[1].Paragraphs[0].Text = "Lower";
            table.Rows[2].Cells[2].Paragraphs[0].Text = "LowerTail";

            table.Rows[0].Cells[0].MergeHorizontally(1);
            table.Rows[1].Cells[0].MergeVertically(1);

            Assert.Equal(2, table.Rows[0].Cells[0].ColumnSpan);
            Assert.Equal(2, table.Rows[1].Cells[0].RowSpan);

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfDocument pdf = PdfDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Across", text);
            Assert.Contains("Tall", text);
            Assert.Contains("TopTail", text);
            Assert.Contains("Upper", text);
            Assert.Contains("Lower", text);
        }

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfLinkAnnotation horizontal = Assert.Single(info.LinkAnnotations, link => link.Uri == horizontalUri);
        PdfCore.PdfLinkAnnotation vertical = Assert.Single(info.LinkAnnotations, link => link.Uri == verticalUri);
        Assert.Equal("Column span metadata", horizontal.Contents);
        Assert.Equal("Row span metadata", vertical.Contents);
        Assert.True(horizontal.Width > 110D);
        Assert.True(vertical.Height > 30D);
    }

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
                OfficeIMOPageSize = new PdfCore.PageSize(260, 220),
                OfficeIMOMargins = PdfCore.PageMargins.Uniform(12)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
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
                OfficeIMOPageSize = new PdfCore.PageSize(260, 220),
                OfficeIMOMargins = PdfCore.PageMargins.Uniform(12)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
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
                OfficeIMOPageSize = new PdfCore.PageSize(260, 220),
                OfficeIMOMargins = PdfCore.PageMargins.Uniform(12)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);

        int firstRowHeaderCount = pdf.GetPages()
            .SelectMany(page => page.GetWords())
            .Count(word => word.Text == "SoloHdr");
        Assert.Equal(1, firstRowHeaderCount);
    }

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
                OfficeIMOPageSize = new PdfCore.PageSize(400, 500),
                OfficeIMOMargins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var left = Assert.Single(words, word => word.Text == "LeftTbl");
        var center = Assert.Single(words, word => word.Text == "CenterTbl");
        var right = Assert.Single(words, word => word.Text == "RightTbl");

        Assert.True(center.BoundingBox.Left > left.BoundingBox.Left + 70D);
        Assert.True(right.BoundingBox.Left > center.BoundingBox.Left + 70D);
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
                OfficeIMOPageSize = new PdfCore.PageSize(400, 500),
                OfficeIMOMargins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfDocument pdf = PdfDocument.Open(bytes);
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
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Preferred_Width_And_AutoFit_Style() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableLayoutStyle.docx"));

        WordTable preferred = document.AddTable(1, 2);
        preferred.WidthType = TableWidthUnitValues.Dxa;
        preferred.Width = 2880;
        PdfCore.PdfTableStyle preferredStyle = CreateNativeTableStyleForTest(preferred);

        Assert.Equal(144D, preferredStyle.MaxWidth);
        Assert.Equal(10D, preferredStyle.FontSize);
        Assert.False(preferredStyle.AutoFitColumns);

        WordTable autoFit = document.AddTable(1, 2);
        autoFit.Rows[0].Cells[0].Paragraphs[0].Text = "Short";
        autoFit.Rows[0].Cells[1].Paragraphs[0].Text = "Much wider auto fit text";
        autoFit.AutoFitToContents();
        PdfCore.PdfTableStyle autoFitStyle = CreateNativeTableStyleForTest(autoFit);

        Assert.True(autoFitStyle.AutoFitColumns);
        Assert.Null(autoFitStyle.MaxWidth);

        WordTable spaced = document.AddTable(1, 2);
        spaced.StyleDetails!.CellSpacing = 240;
        PdfCore.PdfTableStyle spacedStyle = CreateNativeTableStyleForTest(spaced);

        Assert.Equal(12D, spacedStyle.CellSpacing);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Row_Break_Policies() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableRowBreakPolicies.docx"));
        WordTable table = document.AddTable(2, 1);
        table.Rows[0].AllowRowToBreakAcrossPages = false;
        table.Rows[1].AllowRowToBreakAcrossPages = true;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.True(style.AllowRowBreakAcrossPages);
        Assert.NotNull(style.RowAllowBreakAcrossPages);
        Assert.False(style.RowAllowBreakAcrossPages![0]);
        Assert.True(style.RowAllowBreakAcrossPages![1]);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Multiple_Header_Rows() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableMultipleHeaderRows.docx"));
        WordTable table = document.AddTable(4, 1);
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
        table.Rows[1].RepeatHeaderRowAtTheTopOfEachPage = true;
        table.Rows[3].RepeatHeaderRowAtTheTopOfEachPage = true;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(2, style.HeaderRowCount);
        Assert.Equal(2, style.RepeatHeaderRowCount);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_First_Row_Style_Without_Repeating() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableFirstRowStyleNoRepeat.docx"));
        WordTable table = document.AddTable(3, 1, WordTableStyle.GridTable1Light);
        table.ConditionalFormattingFirstRow = true;
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = false;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(1, style.HeaderRowCount);
        Assert.Equal(0, style.RepeatHeaderRowCount);
    }

    private static void ConfigureMarginTable(WordTable table, string label) {
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.Width = 2880;
        cell.WidthType = TableWidthUnitValues.Dxa;
        cell.Paragraphs[0].Text = label;
    }

    private static void ConfigurePlacementTable(WordTable table, string label, TableRowAlignmentValues alignment) {
        table.Alignment = alignment;
        foreach (WordTableCell cell in table.Rows[0].Cells) {
            cell.Width = 1440;
            cell.WidthType = TableWidthUnitValues.Dxa;
        }

        table.Rows[0].Cells[0].Paragraphs[0].Text = label;
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Value";
    }

    private static void ConfigureCellSpacingTable(WordTable table, string left, string right) {
        foreach (WordTableCell cell in table.Rows[0].Cells) {
            cell.Width = 1440;
            cell.WidthType = TableWidthUnitValues.Dxa;
        }

        table.Rows[0].Cells[0].Paragraphs[0].Text = left;
        table.Rows[0].Cells[1].Paragraphs[0].Text = right;
    }

    private static PdfCore.PdfTableStyle CreateNativeTableStyleForTest(WordTable table) {
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeTableStyle", BindingFlags.NonPublic | BindingFlags.Static)!;
        return Assert.IsType<PdfCore.PdfTableStyle>(method.Invoke(null, new object?[] { table, table.Rows.Count, null }));
    }
}

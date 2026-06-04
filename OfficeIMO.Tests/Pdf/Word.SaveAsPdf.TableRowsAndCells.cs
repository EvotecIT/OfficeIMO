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
                PageSize = new PdfCore.PageSize(400, 500),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
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
                PageSize = new PdfCore.PageSize(320, 260),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
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
        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
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
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
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
                PageSize = new PdfCore.PageSize(360, 260),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
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
        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
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
}

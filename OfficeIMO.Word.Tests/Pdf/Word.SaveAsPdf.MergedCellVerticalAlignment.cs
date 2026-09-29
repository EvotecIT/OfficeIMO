using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_UsesMergedCellAnchorVerticalAlignment() {
        // Word's PDF export ignores alignment declared only on continuation cells.
        double defaultY = RenderMergedCellTextBottom("MergedCellDefault", null, false);
        double continuationY = RenderMergedCellTextBottom("MergedCellContinuation", null, true);
        double anchorCenterY = RenderMergedCellTextBottom("MergedCellAnchorCenter", WordTableVerticalAlignment.Center, true);

        Assert.InRange(Math.Abs(defaultY - continuationY), 0D, 2D);
        Assert.True(defaultY > anchorCenterY + 20D,
            $"Expected anchor alignment to center merged-cell text. Default: {defaultY:0.##}; centered: {anchorCenterY:0.##}.");
    }

    private double RenderMergedCellTextBottom(string name, WordTableVerticalAlignment? anchorAlignment, bool alignContinuations) {
        string docxPath = Path.Combine(_directoryWithFiles, name + ".docx");
        string pdfPath = Path.Combine(_directoryWithFiles, name + ".pdf");
        using (WordDocument document = WordDocument.Create(docxPath)) {
            WordTable table = document.AddTable(3, 2);
            foreach (WordTableRow row in table.Rows) {
                row.Height = 560;
                row._tableRow.TableRowProperties!.GetFirstChild<TableRowHeight>()!.HeightType = HeightRuleValues.Exact;
            }

            table.Rows[0].Cells[0].Paragraphs[0].Text = "MergedAnchor";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "FirstPeer";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "SecondPeer";
            table.Rows[2].Cells[1].Paragraphs[0].Text = "ThirdPeer";
            table.Rows[0].Cells[0].MergeVertically(2);

            if (anchorAlignment.HasValue) {
                table.Rows[0].Cells[0].VerticalAlignment = anchorAlignment.Value;
            }
            if (alignContinuations) {
                table.Rows[1].Cells[0].VerticalAlignment = WordTableVerticalAlignment.Bottom;
                table.Rows[2].Cells[0].VerticalAlignment = WordTableVerticalAlignment.Center;
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
        }

        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        return Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "MergedAnchor").BoundingBox.Bottom;
    }
}

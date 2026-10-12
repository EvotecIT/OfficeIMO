using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, WordTableVerticalAlignment.Center)]
    [InlineData(true, WordTableVerticalAlignment.Center)]
    [InlineData(false, WordTableVerticalAlignment.Bottom)]
    [InlineData(true, WordTableVerticalAlignment.Bottom)]
    public void SaveAsPdf_UnspacedNeighborsShareTheirVerticalMarginAlignmentFrame(bool minimum, WordTableVerticalAlignment alignment) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 1, 2, 0);
        WordTableRow row = table.Rows[0];
        if (minimum) row.MinimumHeight = 2400;
        else row.Height = 2400;
        row.Cells[0].MarginTopWidth = 120;
        row.Cells[0].MarginBottomWidth = 0;
        row.Cells[1].MarginTopWidth = 0;
        row.Cells[1].MarginBottomWidth = 180;
        foreach (WordTableCell cell in row.Cells) cell.VerticalAlignment = alignment;
        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Equal(xml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Single(pdf.GetPages());
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.Equal(2, words.Length);
        double first = words.Single(word => word.Text == "Frame0").Letters[0].StartBaseLine.Y;
        double next = words.Single(word => word.Text == "Frame1").Letters[0].StartBaseLine.Y;
        Assert.Equal(first, next, 3);
    }

    [Theory]
    [InlineData(false, 0, 120D)]
    [InlineData(true, 0, 120D)]
    [InlineData(false, 1, 129D)]
    [InlineData(true, 1, 135D)]
    [InlineData(false, 2, 129D)]
    [InlineData(true, 2, 135D)]
    public void SaveAsPdf_UnspacedTableRowHeightIncludesEffectiveMargins(bool minimum, int marginMode, double advance) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 2, 2, 0);
        foreach (WordTableRow row in table.Rows) {
            if (minimum) row.MinimumHeight = 2400;
            else row.Height = 2400;
            row.Cells[0].MarginTopWidth = marginMode == 0 ? (short)0 : (short)120;
            row.Cells[0].MarginBottomWidth = marginMode == 1 ? (short)180 : (short)0;
            row.Cells[1].MarginTopWidth = 0;
            row.Cells[1].MarginBottomWidth = marginMode == 2 ? (short)180 : (short)0;
        }

        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Equal(xml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Single(pdf.GetPages());
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.Equal(4, words.Length);
        double first = words.Single(word => word.Text == "Frame0").Letters[0].StartBaseLine.Y;
        double next = words.Single(word => word.Text == "Frame2").Letters[0].StartBaseLine.Y;
        Assert.Equal(advance, first - next, 3);
        double neighbor = words.Single(word => word.Text == "Frame1").Letters[0].StartBaseLine.Y;
        Assert.Equal(first, neighbor, 3);
    }

    [Theory]
    [InlineData(false, false, false, 129D)]
    [InlineData(true, false, false, 135D)]
    [InlineData(false, true, false, 129D)]
    [InlineData(true, true, false, 135D)]
    [InlineData(false, true, true, 120D)]
    public void SaveAsPdf_VerticalMergeContinuationMarginsContributeToPhysicalRowHeight(
        bool minimum, bool conditional, bool clear, double advance) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 3, 2, 0);
        foreach (WordTableRow row in table.Rows) {
            if (minimum) row.MinimumHeight = 2400;
            else row.Height = 2400;
            foreach (WordTableCell cell in row.Cells) {
                cell.MarginTopWidth = 120;
                cell.MarginBottomWidth = conditional ? null : (short)0;
            }
        }
        if (conditional) {
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new W.Style(new W.BasedOn { Val = "TableGrid" },
                    new W.TableStyleProperties(new W.StyleTableCellProperties(
                        new W.TableCellMargin(new W.BottomMargin { Width = "180", Type = W.TableWidthUnitValues.Dxa }))) {
                        Type = W.TableStyleOverrideValues.FirstColumn
                    }) { Type = W.StyleValues.Table, StyleId = "ContinuationPadding" });
            table._tableProperties!.TableStyle = new W.TableStyle { Val = "ContinuationPadding" };
            table.ConditionalFormattingFirstColumn = true;
            if (clear) table.Rows[1].Cells[0].MarginBottomWidth = 0;
        } else {
            table.Rows[1].Cells[0].MarginBottomWidth = 180;
        }
        table.MergeCells(0, 0, 2, 1);
        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Equal(xml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Single(pdf.GetPages());
        var words = pdf.GetPage(1).GetWords().ToArray();
        double first = words.Single(word => word.Text == "Frame3").Letters[0].StartBaseLine.Y;
        double next = words.Single(word => word.Text == "Frame5").Letters[0].StartBaseLine.Y;
        Assert.Equal(advance, first - next, 3);
    }

    [Theory]
    [InlineData(0, 54.5D)]
    [InlineData(180, 63.5D)]
    public void SaveAsPdf_SpacedVerticalMergeContinuationIncludesBottomMargin(short bottom, double advance) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 3, 2, 120);
        foreach (WordTableRow row in table.Rows) row.Height = 960;
        table.Rows[1].Cells[0].MarginBottomWidth = bottom;
        table.MergeCells(0, 0, 2, 1);
        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Single(pdf.GetPages());
        var words = pdf.GetPage(1).GetWords().ToArray();
        double first = words.Single(word => word.Text == "Frame3").Letters[0].StartBaseLine.Y;
        double next = words.Single(word => word.Text == "Frame5").Letters[0].StartBaseLine.Y;
        Assert.Equal(advance, first - next, 3);
    }
}

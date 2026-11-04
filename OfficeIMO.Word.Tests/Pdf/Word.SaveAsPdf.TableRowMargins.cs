using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, 129D)]
    [InlineData(true, 136D)]
    public void SaveAsPdf_ZeroSpacingRowsShareMarginsAndRetainTheirHeight(bool minimum, double advance) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 2, 2, 0);
        table.StyleDetails!.CellSpacing = global::OfficeIMO.OfficeLength.Points(0);
        foreach (WordTableRow row in table.Rows) {
            if (minimum) row.MinimumHeight = global::OfficeIMO.OfficeLength.Points(120D);
            else row.Height = global::OfficeIMO.OfficeLength.Points(120D);
            row.Cells[0].MarginTop = global::OfficeIMO.OfficeLength.Points(6D);
            row.Cells[1].MarginBottom = global::OfficeIMO.OfficeLength.Points(9D);
        }
        using var pdf = PdfPigDocument.Open(document.ToPdf(BorderFramePdfOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        double first = Assert.Single(words, word => word.Text == "Frame0").Letters[0].StartBaseLine.Y;
        double neighbor = Assert.Single(words, word => word.Text == "Frame1").Letters[0].StartBaseLine.Y;
        double next = Assert.Single(words, word => word.Text == "Frame2").Letters[0].StartBaseLine.Y;
        Assert.Equal(first, neighbor, 3);
        Assert.Equal(advance, first - next, 3);
    }

    [Theory]
    [InlineData(false, WordTableVerticalAlignment.Center)]
    [InlineData(true, WordTableVerticalAlignment.Center)]
    [InlineData(false, WordTableVerticalAlignment.Bottom)]
    [InlineData(true, WordTableVerticalAlignment.Bottom)]
    public void SaveAsPdf_UnspacedNeighborsShareTheirVerticalMarginAlignmentFrame(bool minimum, WordTableVerticalAlignment alignment) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 1, 2, 0);
        WordTableRow row = table.Rows[0];
        if (minimum) row.MinimumHeight = global::OfficeIMO.OfficeLength.Points(120D);
        else row.Height = global::OfficeIMO.OfficeLength.Points(120D);
        row.Cells[0].MarginTop = global::OfficeIMO.OfficeLength.Points(6D);
        row.Cells[0].MarginBottom = global::OfficeIMO.OfficeLength.Zero;
        row.Cells[1].MarginTop = global::OfficeIMO.OfficeLength.Zero;
        row.Cells[1].MarginBottom = global::OfficeIMO.OfficeLength.Points(9D);
        foreach (WordTableCell cell in row.Cells) cell.VerticalAlignment = alignment;
        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdf(BorderFramePdfOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
    [InlineData(true, 0, 120.5D)]
    [InlineData(false, 1, 129D)]
    [InlineData(true, 1, 135.5D)]
    [InlineData(false, 2, 129D)]
    [InlineData(true, 2, 135.5D)]
    public void SaveAsPdf_UnspacedTableRowHeightIncludesEffectiveMargins(bool minimum, int marginMode, double advance) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 2, 2, 0);
        foreach (WordTableRow row in table.Rows) {
            if (minimum) row.MinimumHeight = global::OfficeIMO.OfficeLength.Points(120D);
            else row.Height = global::OfficeIMO.OfficeLength.Points(120D);
            row.Cells[0].MarginTop = global::OfficeIMO.OfficeLength.Points((marginMode == 0 ? (short)0 : (short)120) / 20D);
            row.Cells[0].MarginBottom = global::OfficeIMO.OfficeLength.Points((marginMode == 1 ? (short)180 : (short)0) / 20D);
            row.Cells[1].MarginTop = global::OfficeIMO.OfficeLength.Zero;
            row.Cells[1].MarginBottom = global::OfficeIMO.OfficeLength.Points((marginMode == 2 ? (short)180 : (short)0) / 20D);
        }

        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdf(BorderFramePdfOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
    [InlineData(true, false, false, 135.5D)]
    [InlineData(false, true, false, 129D)]
    [InlineData(true, true, false, 135.5D)]
    [InlineData(false, true, true, 120D)]
    public void SaveAsPdf_VerticalMergeContinuationMarginsContributeToPhysicalRowHeight(
        bool minimum, bool conditional, bool clear, double advance) {
        // Minimum rows include both half-width TableGrid border insets, as in the unmerged control.
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 3, 2, 0);
        foreach (WordTableRow row in table.Rows) {
            if (minimum) row.MinimumHeight = global::OfficeIMO.OfficeLength.Points(120D);
            else row.Height = global::OfficeIMO.OfficeLength.Points(120D);
            foreach (WordTableCell cell in row.Cells) {
                cell.MarginTop = global::OfficeIMO.OfficeLength.Points(6D);
                cell.MarginBottom= conditional ? null : global::OfficeIMO.OfficeLength.Points(((short)0) / 20D);
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
            if (clear) table.Rows[1].Cells[0].MarginBottom = global::OfficeIMO.OfficeLength.Zero;
        } else {
            table.Rows[1].Cells[0].MarginBottom = global::OfficeIMO.OfficeLength.Points(9D);
        }
        table.MergeCells(0, 0, 2, 1);
        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdf(BorderFramePdfOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
        foreach (WordTableRow row in table.Rows) row.Height = global::OfficeIMO.OfficeLength.Points(48D);
        table.Rows[1].Cells[0].MarginBottom = global::OfficeIMO.OfficeLength.Points((bottom) / 20D);
        table.MergeCells(0, 0, 2, 1);
        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdf(BorderFramePdfOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Single(pdf.GetPages());
        var words = pdf.GetPage(1).GetWords().ToArray();
        double first = words.Single(word => word.Text == "Frame3").Letters[0].StartBaseLine.Y;
        double next = words.Single(word => word.Text == "Frame5").Letters[0].StartBaseLine.Y;
        Assert.Equal(advance, first - next, 3);
    }
}

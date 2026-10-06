using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_HiddenTableParagraphJoinsRetainSourceAndCellBoundaries(bool nativeDoc, bool blockControl) {
        using WordDocument source = CreateJoinedParagraphDocument();
        WordTable table = source.AddTable(1, 2);
        WordTableCell cell = table.Rows[0].Cells[0];
        WordParagraph alpha = cell.AddParagraph("ALPHA", removeExistingParagraphs: true); HideJoinMark(alpha, true);
        WordParagraph blank = cell.AddParagraph("SECRET"); blank.Hidden = true; HideJoinMark(blank, true);
        WordParagraph beta = cell.AddParagraph("BETA"); HideJoinMark(beta, true);
        table.Rows[0].Cells[1].AddParagraph("OTHER", removeExistingParagraphs: true);
        if (blockControl) {
            alpha._paragraph.Remove(); blank._paragraph.Remove(); beta._paragraph.Remove();
            cell._tableCell.Append(new W.SdtBlock(new W.SdtProperties(new W.Tag { Val = "CellJoin" }),
                new W.SdtContentBlock(alpha._paragraph, blank._paragraph, beta._paragraph)));
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.Contains("ALPHABETA", pdf.GetPage(1).Text);
        Assert.DoesNotContain("SECRET", pdf.GetPage(1).Text);
        Assert.Contains("OTHER", pdf.GetPage(1).Text);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(letters.First(letter => letter.Value == "A").StartBaseLine.Y,
            Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_HiddenTableJoinRetainsFirstAlignmentFinalSpacingAndContinuationMetrics(bool nativeDoc) {
        using WordDocument source = CreateJoinedParagraphDocument();
        WordTable table = source.AddTable(1, 1);
        table.WidthType = WordTableWidthUnit.Dxa; table.Width = 9360;
        table.LayoutMode = WordTableLayoutMode.Fixed; table.GridColumnWidth = new List<int> { 9360 };
        WordTableCell cell = table.Rows[0].Cells[0]; cell.WidthType = WordTableWidthUnit.Dxa; cell.Width = 9360;
        WordParagraph alpha = cell.AddParagraph("ALPHA", removeExistingParagraphs: true);
        alpha.FontSize = 16; alpha.ParagraphAlignment = WordParagraphAlignment.Right;
        alpha.LineSpacingAfterPoints = 80; HideJoinMark(alpha, true);
        WordParagraph beta = cell.AddParagraph("BETA"); beta.FontSize = 8;
        beta._run!.Append(new W.Break(), new W.Text("TAIL"), new W.Break(), new W.Text("END"));
        beta.LineSpacingBeforePoints = 80; beta.LineSpacingAfterPoints = 20;
        cell.AddParagraph("SECOND");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = OpenJoinedParagraphPdf(document);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(letters.First(letter => letter.Value == "A").StartBaseLine.Y,
            Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
        Assert.True(letters.First(letter => letter.Value == "A").StartBaseLine.X > 400);
        var words = pdf.GetPage(1).GetWords().ToList();
        var tail = Assert.Single(words, word => word.Text == "TAIL"); var end = Assert.Single(words, word => word.Text == "END");
        Assert.InRange(tail.BoundingBox.Bottom - end.BoundingBox.Bottom, 9D, 10D);
        var second = Assert.Single(words, word => word.Text == "SECOND");
        Assert.InRange(end.BoundingBox.Bottom - second.BoundingBox.Bottom, 32D, 36D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_HiddenTableJoinPreservesNoteOnlyReferences(bool nativeDoc) {
        using WordDocument source = CreateJoinedParagraphDocument();
        WordTableCell cell = source.AddTable(1, 1).Rows[0].Cells[0];
        WordParagraph first = cell.AddParagraph(removeExistingParagraphs: true); HideJoinMark(first, true);
        first.AddFootNote("NOTETEXT"); cell.AddParagraph();
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.Contains("NOTETEXT", pdf.GetPage(1).Text);
        Assert.Equal(2, pdf.GetPage(1).Letters.Count(letter => letter.Value == "1"));
    }

    [Fact]
    public void SaveAsPdf_HiddenTableJoinRetainsUnsupportedObjectBoundaryWithDiagnostic() {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        WordParagraph alpha = cell.AddParagraph("ALPHA", removeExistingParagraphs: true); HideJoinMark(alpha, true);
        alpha._paragraph.ParagraphProperties!.Shading = new W.Shading { Fill = "FFFF00" };
        cell.AddParagraph("BETA");
        var result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeHiddenParagraphJoinUnsupported" && warning.Source == "table cell");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.Contains("ALPHA", pdf.GetPage(1).Text); Assert.Contains("BETA", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_UnstyledTableUsesDocumentDefaultCellMargins(bool nativeDoc) {
        using WordDocument source = CreateJoinedParagraphDocument();
        WordTable table = source.AddTable(1, 1); table._tableProperties!.TableStyle = null;
        table.Rows[0].Cells[0].AddParagraph("MARGIN", removeExistingParagraphs: true);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.InRange(pdf.GetPage(1).Letters.First(letter => letter.Value == "M").StartBaseLine.X, 77.3D, 77.5D);
    }

    [Fact]
    public void SaveAsPdf_UnstyledTableRetainsExplicitPdfDefaultPadding() {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordTable table = document.AddTable(1, 1); table._tableProperties!.TableStyle = null;
        table.Rows[0].Cells[0].AddParagraph("MARGIN", removeExistingParagraphs: true);
        var options = new WordToPdfOptions { IncludePageNumbers = false,
            PdfOptions = new PdfOptions { DefaultTableStyle = new PdfTableStyle { CellPaddingLeft = 20D } } };
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(options));
        Assert.InRange(pdf.GetPage(1).Letters.First(letter => letter.Value == "M").StartBaseLine.X, 91.9D, 92.1D);
    }

    [Fact]
    public void SaveAsPdf_HiddenTableJoinPreservesUnresolvedConfiguredTypographyWithDiagnostic() {
        using WordDocument document = CreateJoinedParagraphDocument();
        W.Style normal = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleRunProperties = null;
        WordTable table = document.AddTable(1, 1); table._tableProperties!.TableStyle = null;
        WordParagraph alpha = table.Rows[0].Cells[0].AddParagraph("ALPHA", removeExistingParagraphs: true); HideJoinMark(alpha, true);
        table.Rows[0].Cells[0].AddParagraph("BETA");
        var result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false,
            PdfOptions = new PdfOptions { DefaultTableStyle = new PdfTableStyle { FontSize = 18D } } });
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeHiddenParagraphJoinUnsupported");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.Equal(18D, pdf.GetPage(1).Letters.First(letter => letter.Value == "A").FontSize, 3);
        Assert.Equal(18D, Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "B").FontSize, 3);
    }

    [Fact]
    public void LegacyDoc_TablePartialDefaultMarginsRetainHorizontalInsetAndExplicitZeroCellOverride() {
        using WordDocument source = CreateJoinedParagraphDocument();
        WordTable table = source.AddTable(1, 2);
        table.StyleDetails!.MarginDefaultTopWidth = 60;
        table._tableProperties!.TableStyle = null;
        table.Rows[0].Cells[0].AddParagraph("DEFAULT", removeExistingParagraphs: true);
        WordTableCell second = table.Rows[0].Cells[1]; second.AddParagraph("ZERO", removeExistingParagraphs: true);
        second.MarginLeftWidth = 0;
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordTableCell first = restored.Tables[0].Rows[0].Cells[0]; second = restored.Tables[0].Rows[0].Cells[1];
        Assert.Equal((short)60, first.MarginTopWidth);
        Assert.Equal((short)108, first.MarginLeftWidth); Assert.Equal((short)108, first.MarginRightWidth);
        Assert.Equal((short)0, second.MarginLeftWidth); Assert.Equal((short)108, second.MarginRightWidth);
    }
}

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
        Assert.Equal(5.4D, CreateNativeTableStyleForTest(document.Tables[0]).CellPaddingLeft);
        double expectedX = nativeDoc ? 72D : 77.4D;
        Assert.InRange(pdf.GetPage(1).Letters.First(letter => letter.Value == "M").StartBaseLine.X, expectedX - .1D, expectedX + .1D);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void SaveAsPdf_HiddenTableJoinRetainsInlinePictureBoundaryAndDiagnostic(bool nativeDoc, bool nested) {
        using WordDocument source = CreateJoinedParagraphDocument();
        WordTable table = source.AddTable(1, 1);
        if (nested) table = table.Rows[0].Cells[0].AddTable(1, 1);
        WordTableCell cell = table.Rows[0].Cells[0];
        WordParagraph alpha = cell.AddParagraph("ALPHA", removeExistingParagraphs: true);
        alpha.AddImage(Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png"), 18, 12);
        HideJoinMark(alpha, true);
        cell.AddParagraph("BETA");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        var result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeHiddenParagraphJoinUnsupported" && warning.Source == "table cell");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        var page = pdf.GetPage(1);
        Assert.Contains("ALPHA", page.Text);
        Assert.Contains("BETA", page.Text);
        Assert.Single(page.GetImages());
        Assert.NotEqual(page.Letters.First(letter => letter.Value == "A").StartBaseLine.Y,
            Assert.Single(page.Letters, letter => letter.Value == "B").StartBaseLine.Y);
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml);
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

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void SaveAsPdf_TableNormalChildPropertiesDoNotAlterNamedOrInheritedLayout(int tableMode) {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordTable table = document.AddTable(1, 1, tableMode < 2 ? WordTableStyle.TableNormal : WordTableStyle.TableGrid);
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        if (tableMode == 1) table._tableProperties!.TableStyle = null;
        if (tableMode == 3) {
            styles.Append(new W.Style(new W.StyleName { Val = "Derived Grid" }, new W.BasedOn { Val = "TableGrid" }) {
                StyleId = "DerivedGrid", Type = W.StyleValues.Table, CustomStyle = true
            });
            table._tableProperties!.TableStyle = new W.TableStyle { Val = "DerivedGrid" };
        }
        table.Rows[0].Cells[0].AddParagraph("MARGIN", removeExistingParagraphs: true);
        using var baseline = OpenJoinedParagraphPdf(document);
        var before = baseline.GetPage(1).Letters.First(letter => letter.Value == "M");
        W.Style normal = styles.Elements<W.Style>().Single(style => style.StyleId?.Value == "TableNormal");
        normal.StyleRunProperties = new W.StyleRunProperties(new W.Color { Val = "FF0000" });
        normal.StyleParagraphProperties = new W.StyleParagraphProperties(new W.SpacingBetweenLines { Before = "180" });
        W.StyleTableProperties properties = normal.GetFirstChild<W.StyleTableProperties>()!;
        properties.Shading = new W.Shading { Fill = "FFFF00" };
        properties.TableCellMarginDefault!.TopMargin = new W.TopMargin { Width = "144", Type = W.TableWidthUnitValues.Dxa };
        using var pdf = OpenJoinedParagraphPdf(document);
        var after = pdf.GetPage(1).Letters.First(letter => letter.Value == "M");
        Assert.Equal(before.StartBaseLine.X, after.StartBaseLine.X, 3);
        Assert.Equal(before.StartBaseLine.Y, after.StartBaseLine.Y, 3);
        Assert.Equal(before.Color.ToRGBValues(), after.Color.ToRGBValues());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void SaveAsPdf_TablePartialMarginsRetainMissingSidesAndExplicitOverrides(bool nativeDoc, bool configuredPdf) {
        using WordDocument source = CreateJoinedParagraphDocument();
        WordTable table = source.AddTable(1, 2);
        table.WidthType = WordTableWidthUnit.Dxa; table.Width = 9360;
        table.LayoutMode = WordTableLayoutMode.Fixed; table.GridColumnWidth = new List<int> { 4680, 4680 };
        foreach (WordTableCell cell in table.Rows[0].Cells) { cell.WidthType = WordTableWidthUnit.Dxa; cell.Width = 4680; }
        table.StyleDetails!.MarginDefaultTopWidth = 60;
        table._tableProperties!.TableStyle = null;
        table.Rows[0].Cells[0].AddParagraph("MARGIN", removeExistingParagraphs: true);
        table.Rows[0].Cells[1].AddParagraph("ZERO", removeExistingParagraphs: true);
        table.Rows[0].Cells[1].MarginLeftWidth = 0;
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        var options = new WordToPdfOptions { IncludePageNumbers = false };
        if (configuredPdf) options.PdfOptions = new PdfOptions { DefaultTableStyle = new PdfTableStyle { CellPaddingLeft = 20D } };
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(options));
        var margin = pdf.GetPage(1).Letters.First(letter => letter.Value == "M");
        double expectedX = configuredPdf ? 92D : nativeDoc ? 72D : 77.4D;
        Assert.InRange(margin.StartBaseLine.X, expectedX - .1D, expectedX + .1D);
        // Per-cell zero overrides both Word's missing-side defaults and the configured PDF padding.
        double expectedZeroX = nativeDoc && !configuredPdf ? 300.6D : 306D;
        Assert.InRange(pdf.GetPage(1).Letters.Single(letter => letter.Value == "Z").StartBaseLine.X, expectedZeroX - .1D, expectedZeroX + .1D);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SaveAsPdf_UnnamedTableRetainsCustomDocumentDefaultMargins(bool nativeDoc, bool configuredPdf) {
        using WordDocument source = CreateJoinedParagraphDocument();
        WordTable table = source.AddTable(1, 1); table._tableProperties!.TableStyle = null;
        table.Rows[0].Cells[0].AddParagraph("MARGIN", removeExistingParagraphs: true);
        W.Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        foreach (W.Style style in styles.Elements<W.Style>().Where(style => style.Type?.Value == W.StyleValues.Table)) style.Default = false;
        styles.Append(new W.Style(new W.StyleName { Val = "Custom Default Padding" }, new W.BasedOn { Val = "TableNormal" },
            new W.StyleTableProperties(new W.TableCellMarginDefault(new W.TopMargin { Width = "60", Type = W.TableWidthUnitValues.Dxa },
                new W.TableCellLeftMargin { Width = 240, Type = W.TableWidthValues.Dxa }, new W.BottomMargin { Width = "0", Type = W.TableWidthUnitValues.Dxa },
                new W.TableCellRightMargin { Width = 180, Type = W.TableWidthValues.Dxa }))) {
                StyleId = "CustomDefaultPadding", Type = W.StyleValues.Table, CustomStyle = true, Default = true });
        string before = source._wordprocessingDocument.MainDocumentPart.Document.OuterXml;
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        Assert.Equal(before, source._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        if (nativeDoc) {
            WordTableCell cell = document.Tables[0].Rows[0].Cells[0];
            Assert.Equal((short)240, cell.MarginLeftWidth); Assert.Equal((short)180, cell.MarginRightWidth);
            Assert.Equal((short)60, cell.MarginTopWidth);
        }
        var options = new WordToPdfOptions { IncludePageNumbers = false };
        if (configuredPdf) options.PdfOptions = new PdfOptions { DefaultTableStyle = new PdfTableStyle { CellPaddingLeft = 20D } };
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(options));
        double expectedX = nativeDoc && !configuredPdf ? 72D : 84D;
        Assert.InRange(pdf.GetPage(1).Letters.First(letter => letter.Value == "M").StartBaseLine.X, expectedX - .1D, expectedX + .1D);
    }
}

using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(12, false, 1)]
    [InlineData(24, false, 1)]
    [InlineData(12, true, 1)]
    [InlineData(12, false, 2)]
    public void SaveAsPdf_EmptyTableParagraphRetainsItsMarkLineHeight(int fontSize, bool spaces, int paragraphs) {
        // Compare persisted PDF geometry with a visible character using the
        // same paragraph-mark font. Empty and spacing-only marks occupy a line
        // without adding synthetic visible characters to the result.
        double control = GetEmptyTableParagraphAdvance(fontSize, spaces ? "   X" : "X", paragraphs);
        double empty = GetEmptyTableParagraphAdvance(fontSize, spaces ? "   " : string.Empty, paragraphs);
        Assert.Equal(control, empty, 2);
        Assert.True(empty > fontSize * paragraphs);
    }

    [Fact]
    public void SaveAsPdf_EmptyTableParagraphKeepsExplicitVerticalMargins() {
        double zero = GetEmptyTableParagraphAdvance(12, string.Empty, 1);
        double inset = GetEmptyTableParagraphAdvance(12, string.Empty, 1, top: 120, bottom: 180);
        Assert.Equal(15D, inset - zero, 2);
    }

    [Fact]
    public void SaveAsPdf_PartialCellMarginsRetainWordHorizontalDefaultsAndExplicitZero() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 1, WordTableStyle.TableGrid);
        table.ConditionalFormattingFirstRow = false;
        table._tableProperties!.TableCellMarginDefault = new W.TableCellMarginDefault(
            new W.TopMargin { Width = "0", Type = W.TableWidthUnitValues.Dxa },
            new W.BottomMargin { Width = "0", Type = W.TableWidthUnitValues.Dxa });
        for (int row = 0; row < 2; row++) {
            WordParagraph paragraph = table.Rows[row].Cells[0].Paragraphs[0];
            paragraph.Text = row == 0 ? "DefaultInset" : "ExplicitZero";
            paragraph.FontFamily = "Arial";
            paragraph.FontSize = 12;
            paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
        }
        table.Rows[1].Cells[0].MarginLeftWidth = 0;
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        double defaultX = Assert.Single(words, word => word.Text == "DefaultInset").Letters[0].StartBaseLine.X;
        double explicitX = Assert.Single(words, word => word.Text == "ExplicitZero").Letters[0].StartBaseLine.X;
        // The explicit zero inset still clears half of the collapsed 0.5pt grid border.
        Assert.Equal(5.4D - 0.25D, defaultX - explicitX, 2);
    }

    private static double GetEmptyTableParagraphAdvance(int fontSize, string text, int paragraphs, short top = 0, short bottom = 0) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(3, 1, WordTableStyle.TableGrid);
        table.ConditionalFormattingFirstRow = false;
        foreach (WordTableRow row in table.Rows) {
            WordTableCell cell = row.Cells[0];
            cell.MarginTopWidth = 0;
            cell.MarginBottomWidth = 0;
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.FontFamily = "Arial";
            paragraph.FontSize = 12;
            paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
            paragraph.LineSpacing = 240;
        }
        table.Rows[0].Cells[0].Paragraphs[0].Text = "TopMarker";
        table.Rows[2].Cells[0].Paragraphs[0].Text = "BottomMarker";
        WordTableCell middle = table.Rows[1].Cells[0];
        middle.MarginTopWidth = top;
        middle.MarginBottomWidth = bottom;
        for (int index = 0; index < paragraphs; index++) {
            WordParagraph paragraph = index == 0 ? middle.Paragraphs[0] : middle.AddParagraph();
            paragraph.FontFamily = "Arial";
            paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
            paragraph.LineSpacing = 240;
            paragraph._paragraph.RemoveAllChildren<W.Run>();
            paragraph._paragraph.ParagraphProperties!.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                new W.FontSize { Val = (fontSize * 2).ToString(System.Globalization.CultureInfo.InvariantCulture) });
            if (text.Length > 0)
                paragraph._paragraph.Append(new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                    new W.FontSize { Val = (fontSize * 2).ToString(System.Globalization.CultureInfo.InvariantCulture) }),
                    new W.Text(text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
        }
        string xml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var words = pdf.GetPage(1).GetWords().ToArray();
        if (string.IsNullOrWhiteSpace(text)) Assert.DoesNotContain(words, word => word.Text == "X");
        return Assert.Single(words, word => word.Text == "TopMarker").Letters[0].StartBaseLine.Y -
            Assert.Single(words, word => word.Text == "BottomMarker").Letters[0].StartBaseLine.Y;
    }
}

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
    public void SaveAsPdf_EmptyParagraphUsesVisibleMarkAndIgnoresHiddenRunSize(bool nativeDoc, bool table) {
        double baseline = EmptyMarkGap(nativeDoc, table, "none");
        foreach (string kind in new[] { "empty", "hidden-default", "hidden-large", "hidden-mark", "hidden-style", "visible-mark-style" }) {
            double expected = baseline * (kind is "hidden-mark" or "hidden-style" ? 1D : 2D);
            double actual = EmptyMarkGap(nativeDoc, table, kind);
            Assert.True(Math.Abs(expected - actual) < 0.01D, $"{kind}: expected {expected}, actual {actual}");
        }
    }

    private static double EmptyMarkGap(bool nativeDoc, bool table, string kind) {
        using WordDocument source = WordDocument.Create();
        W.Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        W.Style normal = styles.Elements<W.Style>().Single(s => s.StyleId?.Value == "Normal");
        normal.StyleRunProperties = new W.StyleRunProperties(
            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
            new W.FontSize { Val = "24" }, new W.FontSizeComplexScript { Val = "24" });
        normal.StyleParagraphProperties = new W.StyleParagraphProperties(new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto
        });
        styles.Append(new W.Style {
            StyleId = "HiddenBlank", Type = W.StyleValues.Paragraph, CustomStyle = true,
            StyleName = new W.StyleName { Val = "Hidden Blank" }, BasedOn = new W.BasedOn { Val = "Normal" },
            StyleRunProperties = new W.StyleRunProperties(new W.Vanish())
        });
        WordTableCell? cell = table ? source.AddTable(1, 1).Rows[0].Cells[0] : null;
        if (cell == null) source.AddParagraph("A"); else cell.AddParagraph("A", removeExistingParagraphs: true);
        if (kind != "none") {
            WordParagraph blank = cell == null ? source.AddParagraph() : cell.AddParagraph();
            blank._paragraph.RemoveAllChildren<W.Run>();
            if (kind != "empty") {
                var run = new W.Run(new W.Text("SECRET"));
                if (!kind.EndsWith("style", StringComparison.Ordinal)) run.RunProperties = new W.RunProperties(new W.Vanish());
                if (kind == "hidden-large") run.RunProperties!.FontSize = new W.FontSize { Val = "96" };
                blank._paragraph.Append(run);
            }
            if (kind.EndsWith("style", StringComparison.Ordinal)) blank.SetStyleId("HiddenBlank");
            if (kind is "hidden-mark" or "visible-mark-style") {
                blank._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
                blank._paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties(
                    new W.Vanish { Val = kind == "hidden-mark" });
            }
        }
        if (cell == null) source.AddParagraph("B"); else cell.AddParagraph("B");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var letters = pdf.GetPage(1).Letters;
        Assert.DoesNotContain("SECRET", pdf.GetPage(1).Text);
        return Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.Y -
            Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y;
    }
}

using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    public void SaveAsPdf_NestedTableWithConfiguredShrinkRetainsTextAndReportsEffectiveFallback(int noteKind, bool sourceGridStyle) {
        using WordDocument document = WordDocument.Create();
        WordParagraph anchor = document.AddParagraph("SHRINKREFERENCE");
        WordParagraph story = anchor;
        if (noteKind == 1) story = anchor.AddFootNote("SHRINKNOTESTART").Paragraph.FootNote!.Paragraphs!.First();
        else if (noteKind == 2) story = anchor.AddEndNote("SHRINKNOTESTART").Paragraph.EndNote!.Paragraphs!.First();
        WordTableStyle sourceStyle = sourceGridStyle ? WordTableStyle.TableGrid : WordTableStyle.TableNormal;
        WordTable outer = story.AddTableAfter(1, 1, sourceStyle);
        ConfigureNestedLayoutTable(outer, 6000, 6000);
        WordTableCell cell = outer.Rows[0].Cells[0];
        ConfigureNestedLayoutParagraph(cell.Paragraphs[0], "SHRINKBEFORE");
        WordTable inner = cell.AddTable(1, 2, sourceStyle);
        ConfigureNestedLayoutTable(inner, 6000, 3000, 3000);
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[0].Paragraphs[0], "SHRINKLEFT");
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[1].Paragraphs[0], "SHRINKRIGHT");
        ConfigureNestedLayoutParagraph(cell.AddParagraph(), "SHRINKAFTER");
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string SourceXml() => main.Document.OuterXml + main.StyleDefinitionsPart?.Styles?.OuterXml +
            main.DocumentSettingsPart?.Settings?.OuterXml + main.FootnotesPart?.Footnotes?.OuterXml +
            main.EndnotesPart?.Endnotes?.OuterXml;
        string original = SourceXml();
        var configuredStyle = new PdfTableStyle { HeaderRowCount = 0, ShrinkTextToFit = true };
        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = new PdfCreateOptions { DefaultTableStyle = configuredStyle }
        });
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        foreach (string token in new[] { "SHRINKBEFORE", "SHRINKLEFT", "SHRINKRIGHT", "SHRINKAFTER" })
            Assert.Single(words, word => word.Text == token);
        if (sourceGridStyle) {
            // A concrete Word table style owns its supported geometry; an
            // unused PDF default must not force its cells into fallback.
            Assert.DoesNotContain(result.Warnings, warning => warning.Code == "native-nested-table-layout-approximated");
        } else {
            string expectedWarning = noteKind == 0 ? "native-nested-table-layout-approximated" :
                noteKind == 1 ? "native-footnote-placement-approximated" : "native-endnote-content-approximated";
            PdfConversionWarning warning = Assert.Single(result.Warnings, warning => warning.Code == expectedWarning);
            Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
            Assert.True(result.HasLoss);
        }
        Assert.True(configuredStyle.ShrinkTextToFit);
        Assert.Equal(original, SourceXml());
    }
}

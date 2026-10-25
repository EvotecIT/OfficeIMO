using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.Rtf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfWordTableCompatibilityGeometryTests {
    [Theory]
    [InlineData(0)]
    [InlineData(300)]
    public void ModernImportPreservesNativeLegacyFixtureContentOrigin(int gapTwips) {
        string name = "table-cell-minimum-exact-gap" + gapTwips.ToString(System.Globalization.CultureInfo.InvariantCulture);
        string directory = Path.Combine(AppContext.BaseDirectory, "Documents", "RtfTableGapNative");
        RtfDocument source = RtfDocument.Load(Path.Combine(directory, name + ".rtf"));
        using WordDocument mapped = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(mapped.ToBytes()));
        using WordDocument native = WordDocument.Load(Path.Combine(directory, name + ".native.docx"));
        Assert.Equal(WordCompatibilityMode.Word2013, reopened.CompatibilitySettings.CompatibilityMode);
        Assert.Equal(WordCompatibilityMode.Word2003, native.CompatibilitySettings.CompatibilityMode);

        // The independently Word-imported DOCX retains legacy layout. Rendering
        // both artifacts through the same owner/resources isolates the bridge's
        // coordinate conversion from different font or PDF producer metrics.
        using PdfPigDocument expected = PdfPigDocument.Open(native.ToPdfBytes(Options()));
        using PdfPigDocument actual = PdfPigDocument.Open(reopened.ToPdfBytes(Options()));
        Assert.Equal(1, expected.NumberOfPages);
        Assert.Equal(expected.NumberOfPages, actual.NumberOfPages);
        foreach (string marker in new[] { "TABLE-BEFORE", "TABLE-MIN-A", "TABLE-EXACT-A", "TABLE-AFTER" }) {
            Assert.InRange(Math.Abs(MarkerX(actual, marker) - MarkerX(expected, marker)), 0D, .03D);
        }
        string expectedText = new string(expected.GetPages().SelectMany(page => page.Letters)
            .SelectMany(letter => letter.Value).Where(character => !char.IsWhiteSpace(character)).ToArray());
        string actualText = new string(actual.GetPages().SelectMany(page => page.Letters)
            .SelectMany(letter => letter.Value).Where(character => !char.IsWhiteSpace(character)).ToArray());
        Assert.Equal(expectedText, actualText);
    }

    [Theory]
    [InlineData(WordCompatibilityMode.None, 720)]
    [InlineData(WordCompatibilityMode.Word2003, 720)]
    [InlineData(WordCompatibilityMode.Word2013, 1020)]
    public void ExportUsesSourceCompatibilityAndPreservesLeftAlignedContentOrigin(
        WordCompatibilityMode mode, int expectedRowIndent) {
        RtfDocument source = RtfDocument.Create();
        source.Settings.SetDefaultFont(source.AddFont("Courier New"));
        source.AddParagraph("BODY");
        RtfTableRow row = source.AddTable(1, 1).Rows[0];
        row.SetCellGap(300).SetLeftIndent(0);
        row.Cells[0].AddParagraph("CELL");
        using WordDocument original = source.ToWordDocument();
        original.CompatibilitySettings.CompatibilityMode = mode;
        WordTable table = Assert.Single(original.Tables);
        table._tableProperties!.TableIndentation = new W.TableIndentation { Width = 1020, Type = W.TableWidthUnitValues.Dxa };

        RtfDocument exported = original.ToRtfDocument();
        RtfTableRow exportedRow = Assert.Single(Assert.Single(exported.Blocks.OfType<RtfTable>()).Rows);
        Assert.Equal(300, exportedRow.CellGapTwips);
        Assert.Equal(expectedRowIndent, exportedRow.LeftIndentTwips);
        Assert.Equal(mode, original.CompatibilitySettings.CompatibilityMode);
        Assert.Equal(1020, table._tableProperties.TableIndentation.Width?.Value);
        using WordDocument returned = exported.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(returned.ToBytes()));
        using PdfPigDocument expected = PdfPigDocument.Open(original.ToPdfBytes(Options()));
        using PdfPigDocument actual = PdfPigDocument.Open(reopened.ToPdfBytes(Options()));
        foreach (string marker in new[] { "BODY", "CELL" }) {
            Assert.InRange(Math.Abs(MarkerX(actual, marker) - MarkerX(expected, marker)), 0D, .03D);
        }
    }

    [Theory]
    [InlineData(0, 720)]
    [InlineData(32767, 720)]
    [InlineData(300, -720)]
    [InlineData(300, int.MaxValue)]
    public void ModernImportAndExportPreserveSupportedGapAndSourceRowIndent(int gapTwips, int rowIndent) {
        RtfDocument source = RtfDocument.Create();
        RtfTableRow row = source.AddTable(1, 1).Rows[0];
        row.SetCellGap(gapTwips).SetLeftIndent(rowIndent);
        row.Cells[0].AddParagraph("CONTENT");

        RtfConversionResult<WordDocument> result = source.ToWordDocumentResult();
        using WordDocument mapped = result.Value;
        Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == "RtfWordTableCellGapFlattened");
        Assert.Equal(rowIndent, Assert.Single(mapped.Tables)._tableProperties?.TableIndentation?.Width?.Value);
        RtfDocument exported = mapped.ToRtfDocument();
        RtfTableRow exportedRow = Assert.Single(Assert.Single(exported.Blocks.OfType<RtfTable>()).Rows);
        Assert.Equal(gapTwips, exportedRow.CellGapTwips);
        Assert.Equal(rowIndent, exportedRow.LeftIndentTwips);
        Assert.Equal("CONTENT", Assert.Single(exportedRow.Cells[0].Paragraphs).ToPlainText());
        Assert.Equal(gapTwips, row.CellGapTwips);
        Assert.Equal(rowIndent, row.LeftIndentTwips);
    }

    private static double MarkerX(PdfPigDocument document, string marker) =>
        Assert.Single(document.GetPages().SelectMany(page => page.GetWords()), word => word.Text == marker)
            .Letters[0].StartBaseLine.X;

    private static WordToPdfOptions Options() => new() {
        IncludePageNumbers = false,
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
    };
}

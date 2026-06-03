using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageExtractorTests {
    [Fact]
    public void SplitExtractAndMerge_PreserveStatementFixtureReadbackForWrapperPipelines() {
        byte[] source = PdfDocumentRasterVisualBaselineTests.CreateLineItemsTwoPage();

        IReadOnlyList<byte[]> splitPages = PdfPageExtractor.SplitPages(source);

        Assert.Equal(2, splitPages.Count);
        Assert.Equal(1, PdfInspector.Inspect(splitPages[0]).PageCount);
        Assert.Equal(1, PdfInspector.Inspect(splitPages[1]).PageCount);

        PdfLogicalDocument firstSplit = PdfLogicalDocument.Load(splitPages[0]);
        PdfLogicalDocument secondSplit = PdfLogicalDocument.Load(splitPages[1]);

        Assert.Contains(firstSplit.TextBlocks, block => NormalizeExtractedText(block.Text).Contains("Statement#4048", StringComparison.Ordinal));
        Assert.Contains(firstSplit.Tables, table => TableContainsRow(table, "Experientiamnostrum", "31,80PLN", "2", "63,60PLN"));
        Assert.Contains(secondSplit.TextBlocks, block => NormalizeExtractedText(block.Text).Contains("Subtotal", StringComparison.Ordinal));
        Assert.Contains(secondSplit.TextBlocks, block => NormalizeExtractedText(block.Text).Contains("Documentnote:", StringComparison.Ordinal));
        Assert.Contains(secondSplit.Tables, table => TableContainsRow(table, "Total", "6397,62PLN"));

        byte[] extractedReversed = PdfPageExtractor.ExtractPageRanges(source, PdfPageRange.ParseMany("2,1"));
        PdfLogicalDocument extractedLogical = PdfLogicalDocument.Load(extractedReversed);

        Assert.Equal(2, PdfInspector.Inspect(extractedReversed).PageCount);
        Assert.Contains(extractedLogical.Pages[0].TextBlocks, block => NormalizeExtractedText(block.Text).Contains("Subtotal", StringComparison.Ordinal));
        Assert.Contains(extractedLogical.Pages[0].Tables, table => TableContainsRow(table, "Total", "6397,62PLN"));
        Assert.Contains(extractedLogical.Pages[1].TextBlocks, block => NormalizeExtractedText(block.Text).Contains("Statement#4048", StringComparison.Ordinal));
        Assert.Contains(extractedLogical.Pages[1].Tables, table => TableContainsRow(table, "Experientiamnostrum", "31,80PLN", "2", "63,60PLN"));

        byte[] mergedReversed = PdfMerger.Merge(splitPages[1], splitPages[0]);
        PdfLogicalDocument mergedLogical = PdfLogicalDocument.Load(mergedReversed);

        Assert.Equal(2, PdfInspector.Inspect(mergedReversed).PageCount);
        Assert.Contains(mergedLogical.Pages[0].TextBlocks, block => NormalizeExtractedText(block.Text).Contains("Subtotal", StringComparison.Ordinal));
        Assert.Contains(mergedLogical.Pages[0].Tables, table => TableContainsRow(table, "Total", "6397,62PLN"));
        Assert.Contains(mergedLogical.Pages[1].TextBlocks, block => NormalizeExtractedText(block.Text).Contains("Statement#4048", StringComparison.Ordinal));
        Assert.Contains(mergedLogical.Pages[1].Tables, table => TableContainsRow(table, "Experientiamnostrum", "31,80PLN", "2", "63,60PLN"));
    }
}

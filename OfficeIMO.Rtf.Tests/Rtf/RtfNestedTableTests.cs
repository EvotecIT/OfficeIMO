using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Rtf;

public class RtfNestedTableTests {
    private const string WordCompatibleNestedRtf = @"{\rtf1\ansi\deff0{\fonttbl{\f0 Calibri;}}
\trowd\trgaph108\cellx5000
\pard\intbl Outer before\par
\pard\intbl\itap2 Inner A\nestcell{\nonesttables\par}Inner B\nestcell{\nonesttables\par}
\pard\intbl\itap2{\*\nesttableprops\trowd\trgaph108\cellx2500\cellx5000\nestrow}{\nonesttables\par}
\pard\intbl\cell\row}";

    [Fact]
    public void Read_Models_Word_Nested_Table_And_Preserves_Source() {
        RtfReadResult result = RtfDocument.Read(WordCompatibleNestedRtf);

        RtfTable outer = Assert.IsType<RtfTable>(Assert.Single(result.Document.Blocks));
        RtfTableCell outerCell = Assert.Single(Assert.Single(outer.Rows).Cells);
        RtfTable nested = Assert.Single(outerCell.Blocks.OfType<RtfTable>());
        RtfTableRow nestedRow = Assert.Single(nested.Rows);

        Assert.Equal("Outer before", outerCell.Paragraphs[0].ToPlainText());
        Assert.Equal(2, nestedRow.Cells.Count);
        Assert.Equal("Inner A", Assert.Single(nestedRow.Cells[0].Paragraphs).ToPlainText());
        Assert.Equal("Inner B", Assert.Single(nestedRow.Cells[1].Paragraphs).ToPlainText());
        Assert.Equal(2500, nestedRow.Cells[0].RightBoundaryTwips);
        Assert.Equal(WordCompatibleNestedRtf, result.ToRtfLossless());
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF101" && diagnostic.Message.Contains("nesttableprops"));
    }

    [Fact]
    public void Semantic_Writer_RoundTrips_Nested_Table_Structure() {
        RtfDocument document = CreateNestedDocument();

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;
        RtfTable outer = Assert.IsType<RtfTable>(Assert.Single(roundTrip.Blocks));
        RtfTable nested = Assert.Single(Assert.Single(Assert.Single(outer.Rows).Cells).Blocks.OfType<RtfTable>());

        Assert.Contains(@"\itap2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\nestcell", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\nesttableprops", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\nestrow", rtf, StringComparison.Ordinal);
        Assert.Equal(new[] { "Inner A", "Inner B" }, nested.Rows[0].Cells.Select(cell => cell.Paragraphs[0].ToPlainText()));
    }

    [Fact]
    public void Semantic_Writer_RoundTrips_Multiple_Nested_Rows() {
        RtfDocument document = CreateNestedDocument();
        RtfTable nested = Assert.Single(Assert.Single(Assert.Single(Assert.IsType<RtfTable>(document.Blocks[0]).Rows).Cells).Blocks.OfType<RtfTable>());
        RtfTableRow secondRow = nested.AddRow();
        secondRow.AddCell(2400).AddParagraph("Inner C");
        secondRow.AddCell(4800).AddParagraph("Inner D");

        RtfDocument roundTrip = RtfDocument.Read(document.ToRtf(new RtfWriteOptions { IncludeGenerator = false })).Document;
        RtfTable roundTripOuter = Assert.IsType<RtfTable>(Assert.Single(roundTrip.Blocks));
        RtfTable roundTripNested = Assert.Single(Assert.Single(Assert.Single(roundTripOuter.Rows).Cells).Blocks.OfType<RtfTable>());

        Assert.Equal(2, roundTripNested.Rows.Count);
        Assert.Equal(new[] { "Inner A", "Inner B" }, roundTripNested.Rows[0].Cells.Select(cell => cell.Paragraphs[0].ToPlainText()));
        Assert.Equal(new[] { "Inner C", "Inner D" }, roundTripNested.Rows[1].Cells.Select(cell => cell.Paragraphs[0].ToPlainText()));
    }

    [Fact]
    public void Html_Writer_Emits_Semantic_Nested_Tables() {
        string html = CreateNestedDocument().ToHtml();

        Assert.Equal(2, CountOccurrences(html, "<table>"));
        Assert.Contains("Outer before", html, StringComparison.Ordinal);
        Assert.Contains("Inner A", html, StringComparison.Ordinal);
        Assert.Contains("Inner B", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Word_Bridge_Preserves_Nested_Tables_In_Both_Directions() {
        RtfDocument source = CreateNestedDocument();

        RtfConversionResult<WordDocument> toWord = source.ToWordDocumentResult();
        using WordDocument word = toWord.Value;
        WordTable outerWordTable = Assert.Single(word.Tables);
        WordTableCell outerWordCell = Assert.Single(Assert.Single(outerWordTable.Rows).Cells);
        Assert.True(outerWordCell.HasNestedTables);
        Assert.Equal(2, Assert.Single(outerWordCell.NestedTables).Rows[0].Cells.Count);

        RtfDocument roundTrip = word.ToRtfDocument();
        RtfTable outer = Assert.IsType<RtfTable>(Assert.Single(roundTrip.Blocks));
        RtfTable nested = Assert.Single(Assert.Single(Assert.Single(outer.Rows).Cells).Blocks.OfType<RtfTable>());
        Assert.Equal(new[] { "Inner A", "Inner B" }, nested.Rows[0].Cells.Select(cell => cell.Paragraphs[0].ToPlainText()));
    }

    [Fact]
    public void Markdown_Flattens_Nested_Table_With_Explicit_Loss_Report() {
        var options = new RtfToMarkdownOptions();

        string markdown = CreateNestedDocument().ToMarkdown(options);

        Assert.Contains("Outer before", markdown, StringComparison.Ordinal);
        Assert.Contains("Inner A", markdown, StringComparison.Ordinal);
        Assert.Contains("Inner B", markdown, StringComparison.Ordinal);
        Assert.Contains(options.Diagnostics, diagnostic => diagnostic.Code == "RTFMD016");
        Assert.Contains(options.ConversionReport.Diagnostics, diagnostic =>
            diagnostic.Code == "RTFMD016" && diagnostic.Action == RtfConversionAction.Flattened);
        Assert.Throws<RtfConversionLossException>(() => options.ConversionReport.RequireNoLoss());
    }

    [Fact]
    public void Pdf_Flattens_Nested_Table_With_Explicit_Loss_Report() {
        var options = new RtfPdfSaveOptions();

        byte[] pdf = CreateNestedDocument().SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Contains("Outer before", text, StringComparison.Ordinal);
        Assert.Contains("Inner A", text, StringComparison.Ordinal);
        Assert.Contains("Inner B", text, StringComparison.Ordinal);
        Assert.Contains(options.RtfConversionReport.Diagnostics, diagnostic =>
            diagnostic.Code == "NestedTableFlattened" && diagnostic.Action == RtfConversionAction.Flattened);
        Assert.Throws<RtfConversionLossException>(() => options.RtfConversionReport.RequireNoLoss());
    }

    private static RtfDocument CreateNestedDocument() {
        RtfDocument document = RtfDocument.Create();
        RtfTable outer = document.AddTable(1, 1);
        RtfTableCell outerCell = outer.Rows[0].Cells[0];
        outerCell.AddParagraph("Outer before");
        RtfTable nested = outerCell.AddTable(1, 2);
        nested.Rows[0].Cells[0].AddParagraph("Inner A");
        nested.Rows[0].Cells[1].AddParagraph("Inner B");
        return document;
    }

    private static int CountOccurrences(string value, string token) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(token, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += token.Length;
        }

        return count;
    }
}

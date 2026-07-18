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
        Assert.Contains(@"\row" + Environment.NewLine + @"\pard\par", rtf, StringComparison.Ordinal);
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
    public void Semantic_Writer_RoundTrips_Adjacent_Nested_Tables_As_Distinct_Blocks() {
        RtfDocument document = RtfDocument.Create();
        RtfTable outer = document.AddTable(1, 1);
        RtfTableCell outerCell = outer.Rows[0].Cells[0];
        outerCell.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("First nested table");
        outerCell.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Second nested table");

        RtfDocument roundTrip = RtfDocument.Read(document.ToRtf(new RtfWriteOptions { IncludeGenerator = false })).Document;
        RtfTable readOuter = Assert.IsType<RtfTable>(Assert.Single(roundTrip.Blocks));
        RtfTable[] nestedTables = Assert.Single(Assert.Single(readOuter.Rows).Cells).Blocks.OfType<RtfTable>().ToArray();

        Assert.Equal(2, nestedTables.Length);
        Assert.Equal("First nested table", nestedTables[0].Rows[0].Cells[0].Paragraphs[0].ToPlainText());
        Assert.Equal("Second nested table", nestedTables[1].Rows[0].Cells[0].Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void Semantic_Writer_RoundTrips_Three_Table_Levels() {
        RtfDocument document = RtfDocument.Create();
        RtfTable outer = document.AddTable(1, 1);
        outer.Rows[0].Cells[0].AddParagraph("Level 1");
        RtfTable levelTwo = outer.Rows[0].Cells[0].AddTable(1, 1);
        levelTwo.Rows[0].Cells[0].AddParagraph("Level 2");
        RtfTable levelThree = levelTwo.Rows[0].Cells[0].AddTable(1, 1);
        levelThree.Rows[0].Cells[0].AddParagraph("Level 3");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;
        RtfTable readOuter = Assert.IsType<RtfTable>(Assert.Single(roundTrip.Blocks));
        RtfTable readLevelTwo = Assert.Single(Assert.Single(Assert.Single(readOuter.Rows).Cells).Blocks.OfType<RtfTable>());
        RtfTable readLevelThree = Assert.Single(Assert.Single(Assert.Single(readLevelTwo.Rows).Cells).Blocks.OfType<RtfTable>());

        Assert.Contains(@"\itap3", rtf, StringComparison.Ordinal);
        Assert.Equal("Level 1", Assert.Single(readOuter.Rows[0].Cells[0].Paragraphs, paragraph => paragraph.ToPlainText().Length > 0).ToPlainText());
        Assert.Equal("Level 2", Assert.Single(readLevelTwo.Rows[0].Cells[0].Paragraphs, paragraph => paragraph.ToPlainText().Length > 0).ToPlainText());
        Assert.Equal("Level 3", Assert.Single(readLevelThree.Rows[0].Cells[0].Paragraphs, paragraph => paragraph.ToPlainText().Length > 0).ToPlainText());
    }

    [Fact]
    public void Semantic_Writer_Emits_Nested_Table_Note_Exactly_Once() {
        RtfDocument document = CreateNestedDocument();
        RtfTable outer = Assert.IsType<RtfTable>(document.Blocks[0]);
        RtfTable nested = Assert.Single(Assert.Single(Assert.Single(outer.Rows).Cells).Blocks.OfType<RtfTable>());
        RtfNote note = document.AddNote(RtfNoteKind.Footnote);
        note.AddParagraph("Nested note");
        nested.Rows[0].Cells[0].Paragraphs[0].AddNoteReference(note, "1");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;

        Assert.Equal(1, CountOccurrences(rtf, @"{\footnote"));
        Assert.Equal("Nested note", Assert.Single(roundTrip.Notes).ToPlainText());
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
    public void Html_Reader_Preserves_Nested_Table_And_Following_Outer_Cell_Content() {
        const string html = "<table><tr><td><p>Outer before</p><table><tr><td>Inner A</td><td>Inner B</td></tr></table><p>Outer after</p></td></tr></table>";

        RtfDocument document = HtmlConversionDocument.Parse(html).ToRtfDocument();

        RtfTable outer = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        RtfTableCell outerCell = Assert.Single(Assert.Single(outer.Rows).Cells);
        RtfTable nested = Assert.Single(outerCell.Blocks.OfType<RtfTable>());
        Assert.Equal(new[] { "Inner A", "Inner B" }, nested.Rows[0].Cells.Select(cell => string.Join("\n", cell.Paragraphs.Select(paragraph => paragraph.ToPlainText()))));
        Assert.Equal(new[] { "Outer before", "Outer after" }, outerCell.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
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

        RtfConversionResult<string> result = CreateNestedDocument().ToMarkdownResult(options);
        string markdown = result.Value;

        Assert.Contains("Outer before", markdown, StringComparison.Ordinal);
        Assert.Contains("Inner A", markdown, StringComparison.Ordinal);
        Assert.Contains("Inner B", markdown, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RTFMD016");
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "RTFMD016" && diagnostic.Action == RtfConversionAction.Flattened);
        Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
    }

    [Fact]
    public void Pdf_Flattens_Nested_Table_With_Explicit_Loss_Report() {
        var options = new RtfPdfSaveOptions();

        PdfCore.PdfDocumentConversionResult result = CreateNestedDocument().ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Outer before", text, StringComparison.Ordinal);
        Assert.Contains("Inner A", text, StringComparison.Ordinal);
        Assert.Contains("Inner B", text, StringComparison.Ordinal);
        Assert.Contains(result.Warnings, warning =>
            warning.Code == "NestedTableFlattened" && warning.Details["RtfAction"] == nameof(RtfConversionAction.Flattened));
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

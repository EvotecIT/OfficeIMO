using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReaderAndFooterRegressionTests {

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsTJArrays() {
        byte[] bytes = BuildPdfWithTjArraySpacing();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello world", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsSingleQuoteOperator() {
        byte[] bytes = BuildPdfWithSingleQuoteOperator();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Matches("Hello\\s+world", text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsDoubleQuoteOperator() {
        byte[] bytes = BuildPdfWithDoubleQuoteOperator();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Matches("Hello\\s+world", text);
    }

    [Fact]
    public void PdfReadDocument_ExtractText_TreatsDoubleQuoteOperatorAsLineAdvance() {
        byte[] bytes = BuildPdfWithDoubleQuoteLineAdvanceOperator();

        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Matches("First\\r?\\nSecond", text);
        Assert.DoesNotContain("FirstSecond", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_TreatsQuoteOperatorsAsLineAdvance() {
        byte[] singleQuote = BuildSingleStreamPdf("BT\n/F1 12 Tf\n72 720 Td\n(First) Tj\n(Second) '\nET\n");
        byte[] doubleQuote = BuildPdfWithDoubleQuoteLineAdvanceOperator();

        Assert.Matches("First\\r?\\nSecond", PdfTextExtractor.ExtractAllText(singleQuote));
        Assert.Matches("First\\r?\\nSecond", PdfTextExtractor.ExtractAllText(doubleQuote));
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ResetsXForLineAdvanceOperators() {
        byte[] tStar = BuildSingleStreamPdf("BT\n/F1 12 Tf\n14 TL\n72 720 Td\n(First) Tj\nT*\n(Second) Tj\nET\n");
        byte[] td = BuildSingleStreamPdf("BT\n/F1 12 Tf\n72 720 Td\n(First) Tj\n0 -14 Td\n(Second) Tj\nET\n");
        byte[] singleQuote = BuildSingleStreamPdf("BT\n/F1 12 Tf\n72 720 Td\n(First) Tj\n(Second) '\nET\n");
        byte[] doubleQuote = BuildPdfWithDoubleQuoteLineAdvanceOperator();

        AssertSecondLineStartsAtFirstLineX(tStar);
        AssertSecondLineStartsAtFirstLineX(td);
        AssertSecondLineStartsAtFirstLineX(singleQuote);
        AssertSecondLineStartsAtFirstLineX(doubleQuote);
    }

    [Fact]
    public void PdfReadPage_ExtractStructured_HonorsHeaderFooterIgnoreBands() {
        byte[] bytes = BuildSingleStreamPdf("BT\n/F1 12 Tf\n1 0 0 1 72 760 Tm\n(Header line) Tj\n1 0 0 1 72 400 Tm\n(Body line) Tj\n1 0 0 1 72 30 Tm\n(Footer line) Tj\nET\n");

        var page = PdfReadDocument.Open(bytes).Pages[0].ExtractStructured(new PdfTextLayoutOptions {
            IgnoreHeaderHeight = 60,
            IgnoreFooterHeight = 60,
            ForceSingleColumn = true
        });

        Assert.Contains("Body line", page.Lines);
        Assert.DoesNotContain("Header line", page.Lines);
        Assert.DoesNotContain("Footer line", page.Lines);
        Assert.Contains(page.LinesDetailed, line => line.Text == "Body line");
        Assert.DoesNotContain(page.LinesDetailed, line => line.Text == "Header line");
        Assert.DoesNotContain(page.LinesDetailed, line => line.Text == "Footer line");
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsTDTextPositioningAsLineAdvance() {
        byte[] bytes = BuildPdfWithTDTextPositioning();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Matches("First\\s+Second", text);
        Assert.DoesNotContain("FirstSecond", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_DoesNotTreatInitialTdAsLineAdvance() {
        byte[] bytes = BuildPdfWithInitialTdTextPositioning();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("First", text, StringComparison.Ordinal);
        Assert.Contains("Second", text, StringComparison.Ordinal);
        Assert.DoesNotMatch("First\\r?\\nSecond", text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_PreservesRepeatedTDLineAdvances() {
        byte[] bytes = BuildPdfWithRepeatedTDTextPositioning();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Matches("First(?:\\r?\\n){2}Second", text);
    }

    [Fact]
    public void PdfTextExtractor_GetMetadata_ReadsHexUtf16InfoStrings() {
        byte[] bytes = BuildPdfWithHexMetadata("Hello metadata", "OfficeIMO");

        var metadata = PdfTextExtractor.GetMetadata(bytes);

        Assert.Equal("Hello metadata", metadata.Title);
        Assert.Equal("OfficeIMO", metadata.Author);
    }

    [Fact]
    public void PdfReadDocument_Metadata_ReadsHexUtf16InfoStrings() {
        byte[] bytes = BuildPdfWithHexMetadata("Hello metadata", "OfficeIMO");

        var document = PdfReadDocument.Open(bytes);

        Assert.Equal("Hello metadata", document.Metadata.Title);
        Assert.Equal("OfficeIMO", document.Metadata.Author);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsOctalEscapedLiteralStrings() {
        byte[] bytes = BuildSingleStreamPdf("BT\n/F1 12 Tf\n72 720 Td\n(Hello\\040octal\\041) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello octal!", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsLineContinuedLiteralStrings() {
        byte[] bytes = BuildSingleStreamPdf("BT\n/F1 12 Tf\n72 720 Td\n(Hello\\\r\nworld) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Helloworld", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsEscapedLiteralStrings() {
        byte[] bytes = BuildSingleStreamPdf("BT\n/F1 12 Tf\n72 720 Td\n(Hello\\040octal\\041) Tj\n( Line\\nbreak) Tj\nET\n");

        var spans = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans();

        Assert.Contains(spans, span => span.Text == "Hello octal!");
        Assert.Contains(spans, span => span.Text == "Line break");
    }

    [Fact]
    public void PdfTextExtractor_GetMetadata_ReadsOctalEscapedLiteralInfoStrings() {
        byte[] bytes = BuildPdfWithLiteralMetadata("(Hello\\040meta\\041)", "(OfficeIMO\\\r\nTeam)");

        var metadata = PdfTextExtractor.GetMetadata(bytes);

        Assert.Equal("Hello meta!", metadata.Title);
        Assert.Equal("OfficeIMOTeam", metadata.Author);
    }

    [Fact]
    public void PdfReadDocument_Metadata_ReadsOctalEscapedLiteralInfoStrings() {
        byte[] bytes = BuildPdfWithLiteralMetadata("(Hello\\040meta\\041)", "(OfficeIMO\\\r\nTeam)");

        var document = PdfReadDocument.Open(bytes);

        Assert.Equal("Hello meta!", document.Metadata.Title);
        Assert.Equal("OfficeIMOTeam", document.Metadata.Author);
    }

    [Fact]
    public void PdfReadDocument_Metadata_ReadsLiteralStringsContainingStreamSubstrings() {
        byte[] bytes = BuildPdfWithLiteralMetadata("(mainstream title)", "(upstream author)");

        var document = PdfReadDocument.Open(bytes);

        Assert.Equal("mainstream title", document.Metadata.Title);
        Assert.Equal("upstream author", document.Metadata.Author);
    }

    [Fact]
    public void PdfReadDocument_Metadata_ReadsLiteralStringsContainingStandaloneParserKeywords() {
        byte[] bytes = BuildPdfWithLiteralMetadata("(stream title)", "(endobj author)");

        var document = PdfReadDocument.Open(bytes);

        Assert.Equal("stream title", document.Metadata.Title);
        Assert.Equal("endobj author", document.Metadata.Author);
    }

}

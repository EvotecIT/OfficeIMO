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
    public void PdfTextExtractor_ExtractAllText_ReadsPagesWithContentStreamArrays() {
        byte[] bytes = BuildPdfWithContentStreamArray();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello world", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_PreservesTextStateAcrossContentStreamArrays() {
        byte[] bytes = BuildPdfWithSplitTextStateContentStreamArray();

        var spans = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans();

        var span = Assert.Single(spans, item => item.Text == "Split state");
        Assert.Equal(72, span.X, 3);
        Assert.Equal(720, span.Y, 3);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsPagesWithIndirectKidsArrayObjects() {
        byte[] bytes = BuildPdfWithIndirectKidsArrayObject();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello indirect kids", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadDocument_Load_PreservesPagesWithDirectContentStreams() {
        byte[] bytes = BuildPdfWithTwoDirectContentPages();

        var document = PdfReadDocument.Open(bytes);

        Assert.Equal(2, document.Pages.Count);
        Assert.NotEqual(document.Pages[0].ObjectNumber, document.Pages[1].ObjectNumber);
    }

    [Fact]
    public void PdfReadDocument_Load_PreservesPagesWithDistinctReferencedContentArrays() {
        byte[] bytes = BuildPdfWithDistinctReferencedContentArrays();

        var document = PdfReadDocument.Open(bytes);

        Assert.Equal(2, document.Pages.Count);
        Assert.Contains("Shared stream page", document.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Shared stream page", document.Pages[1].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_IgnoresCyclicKidsReferences() {
        byte[] bytes = BuildPdfWithCyclicKidsReferences();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello cyclic kids", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_IgnoresDirectFormCycles() {
        var page = CreatePdfReadPageWithDirectFormCycle();

        var spans = page.GetTextSpans();

        Assert.Empty(spans);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsPagesWithIndirectContentArrayObjects() {
        byte[] bytes = BuildPdfWithIndirectContentArrayObject();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello world", text, StringComparison.Ordinal);
    }


    [Fact]
    public void PdfReadDocument_CollectPages_ReadsIndirectKidsArrayObjects() {
        byte[] bytes = BuildPdfWithIndirectKidsArrayObject();

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        Assert.Contains("Hello indirect kids", doc.Pages[0].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsIndirectContentArrayObjects() {
        byte[] bytes = BuildPdfWithIndirectContentArrayObject();

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        string joinedText = string.Concat(doc.Pages[0].GetTextSpans().Select(s => s.Text));
        Assert.Contains("Hello", joinedText, StringComparison.Ordinal);
        Assert.Contains("world", joinedText, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsPageDictionariesWithInlineComments() {
        byte[] bytes = BuildPdfWithCommentedPageDictionary();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello comments", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsPageDictionariesWithInlineComments() {
        byte[] bytes = BuildPdfWithCommentedPageDictionary();

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello comments", span.Text);
    }

}

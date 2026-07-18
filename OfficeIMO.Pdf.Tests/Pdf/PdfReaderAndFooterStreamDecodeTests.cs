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
    public void PdfTextExtractor_ExtractAllText_ReadsFlateCompressedContentStreams() {
        byte[] bytes = BuildPdfWithFlateCompressedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello flate) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello flate", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsAsciiHexEncodedContentStreams() {
        byte[] bytes = BuildPdfWithAsciiHexEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello hex) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello hex", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsAscii85EncodedContentStreams() {
        byte[] bytes = BuildPdfWithAscii85EncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello ascii85) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello ascii85", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsChainedAscii85AndFlateContentStreams() {
        byte[] bytes = BuildPdfWithAscii85AndFlateEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello chained) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello chained", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsChainedAsciiHexAndFlateContentStreamsWithAliases() {
        byte[] bytes = BuildPdfWithAsciiHexAndFlateEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello aliases) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello aliases", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsRunLengthEncodedContentStreams() {
        byte[] bytes = BuildPdfWithRunLengthEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello runlength) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello runlength", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsChainedAscii85AndRunLengthContentStreamsWithAliases() {
        byte[] bytes = BuildPdfWithAscii85AndRunLengthEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello runlength chain) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello runlength chain", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsLzwEncodedContentStreams() {
        byte[] bytes = BuildPdfWithLzwEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello LZW stream) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello LZW stream", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsLzwEncodedContentStreams() {
        byte[] bytes = BuildPdfWithLzwEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello LZW spans) Tj\nET\n");

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello LZW spans", span.Text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsChainedAscii85AndLzwContentStreamsWithEarlyChangeZero() {
        byte[] bytes = BuildPdfWithAscii85AndLzwEncodedStream(
            "BT\n/F1 12 Tf\n72 720 Td\n(Hello LZW chain with early change zero and enough repeated content for wider codes) Tj\nET\n",
            earlyChange: 0);

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello LZW chain with early change zero", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsLzwStreamsWithPredictorDecodeParms() {
        byte[] bytes = BuildPdfWithLzwPredictorEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello LZW predictor) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello LZW predictor", text, StringComparison.Ordinal);
    }


    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsFlateStreamsWithPredictorDecodeParms() {
        byte[] bytes = BuildPdfWithPredictorEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello predictor) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello predictor", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsFlateStreamsWithPredictorDecodeParms() {
        byte[] bytes = BuildPdfWithPredictorEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello predictor) Tj\nET\n");

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello predictor", span.Text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsChainedFiltersWithDecodeParmsArrays() {
        byte[] bytes = BuildPdfWithAscii85AndPredictorEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello predictor chain) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello predictor chain", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsChainedFiltersWithDecodeParmsArrays() {
        byte[] bytes = BuildPdfWithAscii85AndPredictorEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello predictor chain) Tj\nET\n");

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello predictor chain", span.Text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsIndirectDecodeParmsDictionaries() {
        byte[] bytes = BuildPdfWithIndirectPredictorEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello predictor indirect) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello predictor indirect", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsIndirectDecodeParmsDictionaries() {
        byte[] bytes = BuildPdfWithIndirectPredictorEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello predictor indirect) Tj\nET\n");

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello predictor indirect", span.Text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsIndirectDecodeParmsArrayEntries() {
        byte[] bytes = BuildPdfWithAscii85AndIndirectPredictorEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello predictor indirect chain) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello predictor indirect chain", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsIndirectDecodeParmsArrayEntries() {
        byte[] bytes = BuildPdfWithAscii85AndIndirectPredictorEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello predictor indirect chain) Tj\nET\n");

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello predictor indirect chain", span.Text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsFlateStreamsWithTiffPredictorDecodeParms() {
        byte[] bytes = BuildPdfWithTiffPredictorEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello TIFF predictor) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello TIFF predictor", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsFlateStreamsWithTiffPredictorDecodeParms() {
        byte[] bytes = BuildPdfWithTiffPredictorEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello TIFF predictor) Tj\nET\n");

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello TIFF predictor", span.Text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsIndirectFilterNameObjects() {
        byte[] bytes = BuildPdfWithIndirectFilterNameEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello indirect filter) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello indirect filter", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsIndirectFilterNameObjects() {
        byte[] bytes = BuildPdfWithIndirectFilterNameEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello indirect filter) Tj\nET\n");

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello indirect filter", span.Text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsIndirectFilterAndDecodeParmsArrayObjects() {
        byte[] bytes = BuildPdfWithIndirectFilterAndDecodeParmsArrayObjects("BT\n/F1 12 Tf\n72 720 Td\n(Hello indirect arrays) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello indirect arrays", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsIndirectFilterAndDecodeParmsArrayObjects() {
        byte[] bytes = BuildPdfWithIndirectFilterAndDecodeParmsArrayObjects("BT\n/F1 12 Tf\n72 720 Td\n(Hello indirect arrays) Tj\nET\n");

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello indirect arrays", span.Text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsStreamsWithIndirectLengthObjects() {
        byte[] bytes = BuildPdfWithIndirectLengthStreamContainingEndstreamLiteral();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello endstream marker", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsStreamsWithIndirectLengthObjects() {
        byte[] bytes = BuildPdfWithIndirectLengthStreamContainingEndstreamLiteral();

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello endstream marker", span.Text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsStreamsContainingEndobjLiterals() {
        byte[] bytes = BuildSingleStreamPdf("BT\n/F1 12 Tf\n72 720 Td\n(Hello endobj marker) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello endobj marker", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsStreamsContainingEndobjLiterals() {
        byte[] bytes = BuildSingleStreamPdf("BT\n/F1 12 Tf\n72 720 Td\n(Hello endobj marker) Tj\nET\n");

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello endobj marker", span.Text);
    }
}

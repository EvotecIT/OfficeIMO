using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfReaderAndFooterRegressionTests {
    [Fact]
    public void PdfSyntax_ParseObjects_ReadsBooleanAndNullObjects() {
        byte[] bytes = BuildPdfWithBooleanAndNullObjects();

        var (map, _) = PdfSyntax.ParseObjects(bytes);

        Assert.True(map[3].Value is PdfBoolean boolTrue && boolTrue.Value);
        Assert.True(map[4].Value is PdfBoolean boolFalse && !boolFalse.Value);
        Assert.Same(PdfNull.Instance, map[5].Value);
    }

    [Fact]
    public void PdfSyntax_ParseObjects_ReadsBooleanAndNullDictionaryValues() {
        byte[] bytes = BuildPdfWithBooleanAndNullObjects();

        var (map, _) = PdfSyntax.ParseObjects(bytes);

        var metadata = Assert.IsType<PdfDictionary>(map[6].Value);
        Assert.True(metadata.Get<PdfBoolean>("IsTagged")?.Value);
        Assert.False(metadata.Get<PdfBoolean>("NeedsRendering")?.Value ?? true);
        Assert.IsType<PdfNull>(metadata.Items["OptionalContent"]);

        var flags = Assert.IsType<PdfArray>(metadata.Items["Flags"]);
        Assert.True(Assert.IsType<PdfBoolean>(flags.Items[0]).Value);
        Assert.False(Assert.IsType<PdfBoolean>(flags.Items[1]).Value);
        Assert.IsType<PdfNull>(flags.Items[2]);
    }

    [Fact]
    public void PdfReadPage_GetPageSize_InheritsMediaBoxFromPagesNode() {
        byte[] pdfBytes = BuildPdfWithInheritedMediaBox(500, 700);

        var doc = PdfReadDocument.Load(pdfBytes);

        Assert.Single(doc.Pages);
        var (width, height) = doc.Pages[0].GetPageSize();
        Assert.Equal(500, width);
        Assert.Equal(700, height);
    }

    [Fact]
    public void PdfReadPage_GetPageSize_PrefersCropBoxOverMediaBox() {
        byte[] pdfBytes = BuildPdfWithMediaAndCropBoxes(500, 700, 300, 400);

        var doc = PdfReadDocument.Load(pdfBytes);

        Assert.Single(doc.Pages);
        var (width, height) = doc.Pages[0].GetPageSize();
        Assert.Equal(300, width);
        Assert.Equal(400, height);
    }

    [Fact]
    public void PdfReadPage_GetPageSize_ReadsInheritedIndirectMediaBoxArrays() {
        byte[] pdfBytes = BuildPdfWithInheritedIndirectMediaBox(520, 710);

        var doc = PdfReadDocument.Load(pdfBytes);

        Assert.Single(doc.Pages);
        var (width, height) = doc.Pages[0].GetPageSize();
        Assert.Equal(520, width);
        Assert.Equal(710, height);
    }

    [Fact]
    public void PdfReadPage_GetPageSize_PrefersInheritedIndirectCropBoxArrays() {
        byte[] pdfBytes = BuildPdfWithInheritedIndirectCropBox(520, 710, 320, 410);

        var doc = PdfReadDocument.Load(pdfBytes);

        Assert.Single(doc.Pages);
        var (width, height) = doc.Pages[0].GetPageSize();
        Assert.Equal(320, width);
        Assert.Equal(410, height);
    }

    [Fact]
    public void Footer_UsesConfiguredFooterFont() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.TimesRoman,
            ShowPageNumbers = true,
            FooterFormat = "Footer check",
            FooterFont = PdfStandardFont.HelveticaBold,
            FooterAlign = PdfAlign.Left
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("Body text only."))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var footerLine = pdf.GetPage(1).Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderBy(group => group.Key)
            .First()
            .OrderBy(letter => letter.StartBaseLine.X)
            .ToList();

        string footerText = new string(string.Concat(footerLine.Select(letter => letter.Value)).Where(c => !char.IsWhiteSpace(c)).ToArray());
        Assert.Equal("Footercheck", footerText);
        Assert.Contains(footerLine, letter => letter.FontName != null && letter.FontName.Contains("Helvetica-Bold", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(footerLine, letter => letter.FontName != null && letter.FontName.Contains("Times", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsLibraryGeneratedHexText() {
        byte[] bytes = PdfDoc.Create()
            .Paragraph(p => p.Text("Hello extractor"))
            .ToBytes();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello extractor", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsPagesWithContentStreamArrays() {
        byte[] bytes = BuildPdfWithContentStreamArray();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello world", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsPagesWithIndirectKidsArrayObjects() {
        byte[] bytes = BuildPdfWithIndirectKidsArrayObject();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello indirect kids", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsPagesWithIndirectContentArrayObjects() {
        byte[] bytes = BuildPdfWithIndirectContentArrayObject();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello world", text, StringComparison.Ordinal);
    }

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

        Assert.Contains("Hello world", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsDoubleQuoteOperator() {
        byte[] bytes = BuildPdfWithDoubleQuoteOperator();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello world", text, StringComparison.Ordinal);
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

        var document = PdfReadDocument.Load(bytes);

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
    public void PdfTextExtractor_GetMetadata_ReadsOctalEscapedLiteralInfoStrings() {
        byte[] bytes = BuildPdfWithLiteralMetadata("(Hello\\040meta\\041)", "(OfficeIMO\\\r\nTeam)");

        var metadata = PdfTextExtractor.GetMetadata(bytes);

        Assert.Equal("Hello meta!", metadata.Title);
        Assert.Equal("OfficeIMOTeam", metadata.Author);
    }

    [Fact]
    public void PdfReadDocument_Metadata_ReadsOctalEscapedLiteralInfoStrings() {
        byte[] bytes = BuildPdfWithLiteralMetadata("(Hello\\040meta\\041)", "(OfficeIMO\\\r\nTeam)");

        var document = PdfReadDocument.Load(bytes);

        Assert.Equal("Hello meta!", document.Metadata.Title);
        Assert.Equal("OfficeIMOTeam", document.Metadata.Author);
    }

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
    public void PdfTextExtractor_ExtractAllText_ReadsInlineNestedFormResourceDictionaries() {
        byte[] bytes = BuildPdfWithInlineNestedFormResources();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Inline form", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsNestedFormXObjects() {
        byte[] bytes = BuildPdfWithNestedFormInvocations();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Nested form", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_PreservesInlineFormOrdering() {
        byte[] bytes = BuildPdfWithInlineFormTextOrdering();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Before middle after", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_UsesInheritedResourcesForFormXObjects() {
        byte[] bytes = BuildPdfWithInheritedFormResources();

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Form hello");
        Assert.Equal(110, span.X, 3);
        Assert.Equal(220, span.Y, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_TracksRepeatedFormInvocations() {
        byte[] bytes = BuildPdfWithRepeatedFormInvocations();

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var spans = doc.Pages[0].GetTextSpans().Where(s => s.Text == "Repeated form").OrderBy(s => s.X).ToList();
        Assert.Equal(2, spans.Count);
        Assert.Equal(10, spans[0].X, 3);
        Assert.Equal(110, spans[1].X, 3);
        Assert.Equal(20, spans[0].Y, 3);
        Assert.Equal(20, spans[1].Y, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_TracksNestedFormInvocations() {
        byte[] bytes = BuildPdfWithNestedFormInvocations();

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Nested form");
        Assert.Equal(120, span.X, 3);
        Assert.Equal(232, span.Y, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsInlineNestedFormResourceDictionaries() {
        byte[] bytes = BuildPdfWithInlineNestedFormResources();

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Inline form");
        Assert.Equal(10, span.X, 3);
        Assert.Equal(20, span.Y, 3);
    }

    [Fact]
    public void PdfReadDocument_CollectPages_ReadsIndirectKidsArrayObjects() {
        byte[] bytes = BuildPdfWithIndirectKidsArrayObject();

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        Assert.Contains("Hello indirect kids", doc.Pages[0].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsIndirectContentArrayObjects() {
        byte[] bytes = BuildPdfWithIndirectContentArrayObject();

        var doc = PdfReadDocument.Load(bytes);

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

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello comments", span.Text);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsFormResourcesWithEscapedNames() {
        byte[] bytes = BuildPdfWithFormResourceNameEscapes(dictionaryUsesEscapedName: true, contentUsesEscapedName: false);

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Escaped form", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsFormInvocationsWithEscapedNames() {
        byte[] bytes = BuildPdfWithFormResourceNameEscapes(dictionaryUsesEscapedName: false, contentUsesEscapedName: true);

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Escaped form", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsFormResourcesWithEscapedNames() {
        byte[] bytes = BuildPdfWithFormResourceNameEscapes(dictionaryUsesEscapedName: true, contentUsesEscapedName: false);

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Escaped form");
        Assert.Equal(10, span.X, 3);
        Assert.Equal(20, span.Y, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsFormInvocationsWithEscapedNames() {
        byte[] bytes = BuildPdfWithFormResourceNameEscapes(dictionaryUsesEscapedName: false, contentUsesEscapedName: true);

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Escaped form");
        Assert.Equal(10, span.X, 3);
        Assert.Equal(20, span.Y, 3);
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

        var doc = PdfReadDocument.Load(bytes);

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

        var doc = PdfReadDocument.Load(bytes);

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

        var doc = PdfReadDocument.Load(bytes);

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

        var doc = PdfReadDocument.Load(bytes);

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

        var doc = PdfReadDocument.Load(bytes);

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

        var doc = PdfReadDocument.Load(bytes);

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

        var doc = PdfReadDocument.Load(bytes);

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

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello endstream marker", span.Text);
    }

    private static byte[] BuildPdfWithInheritedMediaBox(int width, int height) {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hi) Tj\nET\n";
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            $"<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 {width} {height}] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithBooleanAndNullObjects() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Metadata 6 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [ ] >>",
            "endobj",
            "3 0 obj",
            "true",
            "endobj",
            "4 0 obj",
            "false",
            "endobj",
            "5 0 obj",
            "null",
            "endobj",
            "6 0 obj",
            "<< /IsTagged true /NeedsRendering false /OptionalContent null /Flags [true false null] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithMediaAndCropBoxes(int mediaWidth, int mediaHeight, int cropWidth, int cropHeight) {
        const string streamContent = "BT\n/F1 12 Tf\n72 360 Td\n(Hi) Tj\nET\n";
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {mediaWidth} {mediaHeight}] /CropBox [0 0 {cropWidth} {cropHeight}] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithInheritedIndirectMediaBox(int width, int height) {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hi) Tj\nET\n";
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox 6 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"[0 0 {width} {height}]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithInheritedIndirectCropBox(int mediaWidth, int mediaHeight, int cropWidth, int cropHeight) {
        const string streamContent = "BT\n/F1 12 Tf\n72 360 Td\n(Hi) Tj\nET\n";
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox 6 0 R /CropBox 7 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"[0 0 {mediaWidth} {mediaHeight}]",
            "endobj",
            "7 0 obj",
            $"[0 0 {cropWidth} {cropHeight}]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithContentStreamArray() {
        const string streamOne = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\nET\n";
        const string streamTwo = "BT\n/F1 12 Tf\n72 720 Td\n( world) Tj\nET\n";
        int streamOneLength = Encoding.ASCII.GetByteCount(streamOne);
        int streamTwoLength = Encoding.ASCII.GetByteCount(streamTwo);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents [5 0 R 6 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamOneLength} >>",
            "stream",
            streamOne.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {streamTwoLength} >>",
            "stream",
            streamTwo.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithIndirectKidsArrayObject() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello indirect kids) Tj\nET\n";
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids 6 0 R /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            "[3 0 R]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithIndirectContentArrayObject() {
        const string streamOne = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\nET\n";
        const string streamTwo = "BT\n/F1 12 Tf\n72 720 Td\n( world) Tj\nET\n";
        int streamOneLength = Encoding.ASCII.GetByteCount(streamOne);
        int streamTwoLength = Encoding.ASCII.GetByteCount(streamTwo);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents 7 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamOneLength} >>",
            "stream",
            streamOne.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {streamTwoLength} >>",
            "stream",
            streamTwo.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "[5 0 R 6 0 R]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithCommentedPageDictionary() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello comments) Tj\nET\n";
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page",
            "/Parent 2 0 R",
            "/Resources % resources comment",
            "<< /Font << /F1 4 0 R >> >>",
            "/Contents % contents comment",
            "5 0 R",
            ">>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithFormResourceNameEscapes(bool dictionaryUsesEscapedName, bool contentUsesEscapedName) {
        const string formContent = "BT\n/F1 12 Tf\n10 20 Td\n(Escaped form) Tj\nET\n";
        string pageContentName = contentUsesEscapedName ? "/Fm#31" : "/Fm1";
        string resourceName = dictionaryUsesEscapedName ? "/Fm#31" : "/Fm1";
        string pageContent = $"q\n{pageContentName} Do\nQ\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            $"<< /Type /Page /Parent 2 0 R /Resources << /XObject << {resourceName} 5 0 R >> >> /Contents 6 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /Font << /F1 4 0 R >> >> /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        return BuildSingleStreamPdf(compressedBytes, $"/Filter /FlateDecode /DecodeParms << /Predictor 12 /Columns {streamBytes.Length} >>");
    }

    private static byte[] BuildPdfWithAscii85AndPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(compressedBytes));
        return BuildSingleStreamPdf(encodedBytes, $"/Filter [/ASCII85Decode /FlateDecode] /DecodeParms [null << /Predictor 12 /Columns {streamBytes.Length} >>]");
    }

    private static byte[] BuildPdfWithIndirectPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        return BuildSingleStreamPdfWithExtraObjects(
            compressedBytes,
            "/Filter /FlateDecode /DecodeParms 6 0 R",
            "6 0 obj",
            $"<< /Predictor 12 /Columns {streamBytes.Length} >>",
            "endobj");
    }

    private static byte[] BuildPdfWithAscii85AndIndirectPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(compressedBytes));
        return BuildSingleStreamPdfWithExtraObjects(
            encodedBytes,
            "/Filter [/ASCII85Decode /FlateDecode] /DecodeParms [null 6 0 R]",
            "6 0 obj",
            $"<< /Predictor 12 /Columns {streamBytes.Length} >>",
            "endobj");
    }

    private static byte[] BuildPdfWithTiffPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeTiffPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        return BuildSingleStreamPdf(compressedBytes, $"/Filter /FlateDecode /DecodeParms << /Predictor 2 /Columns {streamBytes.Length} >>");
    }

    private static byte[] BuildPdfWithIndirectFilterNameEncodedStream(string streamContent) {
        byte[] compressedBytes = CompressWithDeflate(Encoding.ASCII.GetBytes(streamContent));
        return BuildSingleStreamPdfWithExtraObjects(
            compressedBytes,
            "/Filter 6 0 R",
            "6 0 obj",
            "/FlateDecode",
            "endobj");
    }

    private static byte[] BuildPdfWithIndirectFilterAndDecodeParmsArrayObjects(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] compressedBytes = CompressWithDeflate(predictedBytes);
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(compressedBytes));
        return BuildSingleStreamPdfWithExtraObjects(
            encodedBytes,
            "/Filter 6 0 R /DecodeParms 7 0 R",
            "6 0 obj",
            "[/ASCII85Decode /FlateDecode]",
            "endobj",
            "7 0 obj",
            "[null 8 0 R]",
            "endobj",
            "8 0 obj",
            $"<< /Predictor 12 /Columns {streamBytes.Length} >>",
            "endobj");
    }

    private static byte[] BuildPdfWithIndirectLengthStreamContainingEndstreamLiteral() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello endstream marker) Tj\nET\n";
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        return BuildSingleStreamPdfWithExtraObjects(
            streamBytes,
            "/Length 6 0 R",
            "6 0 obj",
            streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "endobj");
    }

    private static byte[] BuildPdfWithTjArraySpacing() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n[(Hello) -600 (world)] TJ\nET\n";
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithSingleQuoteOperator() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\n( world) '\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithDoubleQuoteOperator() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\n0 0 ( world) \"\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithFlateCompressedStream(string streamContent) {
        byte[] compressedBytes = CompressWithDeflate(Encoding.ASCII.GetBytes(streamContent));
        return BuildSingleStreamPdf(compressedBytes, "/Filter /FlateDecode");
    }

    private static byte[] BuildPdfWithAsciiHexEncodedStream(string streamContent) {
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAsciiHex(Encoding.ASCII.GetBytes(streamContent)));
        return BuildSingleStreamPdf(encodedBytes, "/Filter /ASCIIHexDecode");
    }

    private static byte[] BuildPdfWithAscii85EncodedStream(string streamContent) {
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(Encoding.ASCII.GetBytes(streamContent)));
        return BuildSingleStreamPdf(encodedBytes, "/Filter /ASCII85Decode");
    }

    private static byte[] BuildPdfWithAscii85AndFlateEncodedStream(string streamContent) {
        byte[] flatedBytes = CompressWithDeflate(Encoding.ASCII.GetBytes(streamContent));
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(flatedBytes));
        return BuildSingleStreamPdf(encodedBytes, "/Filter [/ASCII85Decode /FlateDecode]");
    }

    private static byte[] BuildPdfWithAsciiHexAndFlateEncodedStream(string streamContent) {
        byte[] flatedBytes = CompressWithDeflate(Encoding.ASCII.GetBytes(streamContent));
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAsciiHex(flatedBytes));
        return BuildSingleStreamPdf(encodedBytes, "/Filter [/AHx /Fl]");
    }

    private static byte[] BuildPdfWithRunLengthEncodedStream(string streamContent) {
        byte[] encodedBytes = EncodeRunLength(Encoding.ASCII.GetBytes(streamContent));
        return BuildSingleStreamPdf(encodedBytes, "/Filter /RunLengthDecode");
    }

    private static byte[] BuildPdfWithAscii85AndRunLengthEncodedStream(string streamContent) {
        byte[] runLengthBytes = EncodeRunLength(Encoding.ASCII.GetBytes(streamContent));
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(runLengthBytes));
        return BuildSingleStreamPdf(encodedBytes, "/Filter [/A85 /RL]");
    }

    private static byte[] BuildPdfWithInlineNestedFormResources() {
        const string pageContent = "q\n/Fx Do\nQ\n";
        const string formContent = "BT\n/F1 12 Tf\n10 20 Td\n(Inline form) Tj\nET\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Fx 5 0 R >> >> /Contents 6 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /Font << /F1 4 0 R >> >> /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSingleStreamPdf(string streamContent) {
        return BuildSingleStreamPdf(Encoding.ASCII.GetBytes(streamContent.TrimEnd('\n')));
    }

    private static byte[] BuildSingleStreamPdf(byte[] streamBytes, string extraStreamDictionaryEntries = "") {
        return BuildSingleStreamPdfWithExtraObjects(streamBytes, extraStreamDictionaryEntries);
    }

    private static byte[] BuildSingleStreamPdfWithExtraObjects(byte[] streamBytes, string extraStreamDictionaryEntries = "", params string[] extraObjects) {
        using var ms = new MemoryStream();
        using var writer = new StreamWriter(ms, Encoding.ASCII, 1024, leaveOpen: true);

        writer.WriteLine("%PDF-1.4");
        writer.WriteLine("1 0 obj");
        writer.WriteLine("<< /Type /Catalog /Pages 2 0 R >>");
        writer.WriteLine("endobj");
        writer.WriteLine("2 0 obj");
        writer.WriteLine("<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>");
        writer.WriteLine("endobj");
        writer.WriteLine("3 0 obj");
        writer.WriteLine("<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>");
        writer.WriteLine("endobj");
        writer.WriteLine("4 0 obj");
        writer.WriteLine("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
        writer.WriteLine("endobj");
        writer.WriteLine("5 0 obj");
        writer.Write("<< /Length ");
        writer.Write(streamBytes.Length);
        if (!string.IsNullOrWhiteSpace(extraStreamDictionaryEntries)) {
            writer.Write(' ');
            writer.Write(extraStreamDictionaryEntries.Trim());
        }
        writer.WriteLine(" >>");
        writer.WriteLine("stream");
        writer.Flush();

        ms.Write(streamBytes, 0, streamBytes.Length);

        writer.WriteLine();
        writer.WriteLine("endstream");
        writer.WriteLine("endobj");
        foreach (string extraObject in extraObjects) {
            writer.WriteLine(extraObject);
        }
        writer.WriteLine("trailer");
        writer.WriteLine("<< /Root 1 0 R >>");
        writer.WriteLine("%%EOF");
        writer.Flush();
        return ms.ToArray();
    }

    private static byte[] CompressWithDeflate(byte[] input) {
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(input, 0, input.Length);
        }

        return output.ToArray();
    }

    private static byte[] EncodeUpPredictedRows(byte[] input) {
        var output = new byte[input.Length + 1];
        output[0] = 2;
        Buffer.BlockCopy(input, 0, output, 1, input.Length);
        return output;
    }

    private static byte[] EncodeTiffPredictedRows(byte[] input) {
        if (input.Length == 0) {
            return Array.Empty<byte>();
        }

        var output = new byte[input.Length];
        output[0] = input[0];
        for (int i = 1; i < input.Length; i++) {
            output[i] = unchecked((byte)(input[i] - input[i - 1]));
        }

        return output;
    }

    private static byte[] EncodeRunLength(byte[] input) {
        using var output = new MemoryStream();
        int index = 0;
        while (index < input.Length) {
            int chunkLength = Math.Min(128, input.Length - index);
            output.WriteByte((byte)(chunkLength - 1));
            output.Write(input, index, chunkLength);
            index += chunkLength;
        }

        output.WriteByte(128);
        return output.ToArray();
    }

    private static string EncodeAsciiHex(byte[] input) {
        var sb = new StringBuilder(input.Length * 2 + 1);
        for (int i = 0; i < input.Length; i++) {
            sb.Append(input[i].ToString("X2"));
        }
        sb.Append('>');
        return sb.ToString();
    }

    private static string EncodeAscii85(byte[] input) {
        var sb = new StringBuilder((input.Length * 5 / 4) + 4);
        int index = 0;
        while (index + 4 <= input.Length) {
            uint value =
                ((uint)input[index] << 24) |
                ((uint)input[index + 1] << 16) |
                ((uint)input[index + 2] << 8) |
                input[index + 3];

            if (value == 0) {
                sb.Append('z');
            } else {
                AppendAscii85Tuple(sb, value, 5);
            }

            index += 4;
        }

        int remaining = input.Length - index;
        if (remaining > 0) {
            uint value = 0;
            for (int i = 0; i < remaining; i++) {
                value |= (uint)input[index + i] << (24 - (8 * i));
            }

            AppendAscii85Tuple(sb, value, remaining + 1);
        }

        sb.Append("~>");
        return sb.ToString();
    }

    private static void AppendAscii85Tuple(StringBuilder sb, uint value, int count) {
        char[] encoded = new char[5];
        for (int i = 4; i >= 0; i--) {
            encoded[i] = (char)((value % 85) + '!');
            value /= 85;
        }

        for (int i = 0; i < count; i++) {
            sb.Append(encoded[i]);
        }
    }

    private static byte[] BuildPdfWithHexMetadata(string title, string author) {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [ ] >>",
            "endobj",
            "3 0 obj",
            $"<< /Title <{EncodeUtf16BeHex(title)}> /Author <{EncodeUtf16BeHex(author)}> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 3 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithLiteralMetadata(string titleLiteral, string authorLiteral) {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [ ] >>",
            "endobj",
            "3 0 obj",
            $"<< /Title {titleLiteral} /Author {authorLiteral} >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 3 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithInheritedFormResources() {
        const string pageContent = "q\n/Fx Do\nQ\n";
        const string formContent = "BT\n/F1 12 Tf\n10 20 Td\n(Form hello) Tj\nET\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 7 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Matrix [1 0 0 1 100 200] /Resources 9 0 R /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /XObject 8 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Fx 5 0 R >>",
            "endobj",
            "9 0 obj",
            "<< /Font 10 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /F1 4 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithRepeatedFormInvocations() {
        const string pageContent = "q\n1 0 0 1 0 0 cm\n/Fx Do\nQ\nq\n1 0 0 1 100 0 cm\n/Fx Do\nQ\n";
        const string formContent = "BT\n/F1 12 Tf\n10 20 Td\n(Repeated form) Tj\nET\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 7 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources 9 0 R /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /XObject 8 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Fx 5 0 R >>",
            "endobj",
            "9 0 obj",
            "<< /Font 10 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /F1 4 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithNestedFormInvocations() {
        const string pageContent = "q\n1 0 0 1 100 200 cm\n/FxOuter Do\nQ\n";
        const string outerFormContent = "q\n1 0 0 1 15 25 cm\n/FxInner Do\nQ\n";
        const string innerFormContent = "BT\n/F1 12 Tf\n5 7 Td\n(Nested form) Tj\nET\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int outerFormLength = Encoding.ASCII.GetByteCount(outerFormContent);
        int innerFormLength = Encoding.ASCII.GetByteCount(innerFormContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 7 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 8 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 200 200] /Resources 9 0 R /Length {outerFormLength} >>",
            "stream",
            outerFormContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 50 50] /Resources 11 0 R /Length {innerFormLength} >>",
            "stream",
            innerFormContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /XObject 10 0 R >>",
            "endobj",
            "8 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "9 0 obj",
            "<< /XObject 12 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /FxOuter 5 0 R >>",
            "endobj",
            "11 0 obj",
            "<< /Font 13 0 R >>",
            "endobj",
            "12 0 obj",
            "<< /FxInner 6 0 R >>",
            "endobj",
            "13 0 obj",
            "<< /F1 4 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithInlineFormTextOrdering() {
        const string pageContent = "BT /F1 12 Tf (Before ) Tj /Fx Do ( after) Tj ET\n";
        const string formContent = "BT /F1 12 Tf (middle) Tj ET\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 7 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources 8 0 R /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "5 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Font 9 0 R /XObject 10 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Font 9 0 R >>",
            "endobj",
            "9 0 obj",
            "<< /F1 6 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /Fx 4 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static string EncodeUtf16BeHex(string value) {
        byte[] bom = new byte[] { 0xFE, 0xFF };
        byte[] textBytes = Encoding.BigEndianUnicode.GetBytes(value);
        byte[] bytes = new byte[bom.Length + textBytes.Length];
        Buffer.BlockCopy(bom, 0, bytes, 0, bom.Length);
        Buffer.BlockCopy(textBytes, 0, bytes, bom.Length, textBytes.Length);

        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            sb.Append(bytes[i].ToString("X2"));
        }
        return sb.ToString();
    }
}

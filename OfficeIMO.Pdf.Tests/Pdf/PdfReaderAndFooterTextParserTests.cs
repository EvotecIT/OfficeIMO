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
    public void PdfTextExtractor_ExtractAllText_ReadsLibraryGeneratedHexText() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(p => p.Text("Hello extractor"))
            .ToBytes();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello extractor", text, StringComparison.Ordinal);
    }

    [Fact]
    public void TextContentParser_UsesTwoByteCid_WhenSingleHighByteDecodesToNullGlyph() {
        const string content = "BT /F1 12 Tf 0 0 Td <004400450046> Tj ET";

        var spans = TextContentParser.Parse(
            content,
            static (_, bytes) => DecodeSyntheticCid(bytes),
            static (_, _) => 500);

        Assert.Equal("ABC", string.Concat(spans.Select(static span => span.Text)));

        static string DecodeSyntheticCid(byte[] bytes) {
            if (bytes.Length == 1 && bytes[0] == 0x00) {
                return "\0";
            }

            if (bytes.Length == 2 && bytes[0] == 0x00) {
                if (bytes[1] == 0x44) return "A";
                if (bytes[1] == 0x45) return "B";
                if (bytes[1] == 0x46) return "C";
            }

            return string.Empty;
        }
    }

    [Fact]
    public void TextContentParser_UsesTwoByteCid_WhenWidthEvidenceIdentifiesCidGlyphs() {
        const string content = "BT /F1 12 Tf 0 0 Td <01000101> Tj ET";

        var spans = TextContentParser.Parse(
            content,
            static (_, bytes) => DecodeSyntheticCid(bytes),
            static (_, bytes) => MeasureSyntheticCid(bytes));

        PdfTextSpan span = Assert.Single(spans);
        Assert.Equal("AB", span.Text);
        Assert.Equal(12, span.Advance, 3);

        static string DecodeSyntheticCid(byte[] bytes) {
            if (bytes.Length == 1) {
                return ((char)bytes[0]).ToString();
            }

            if (bytes.Length == 2 && bytes[0] == 0x01) {
                if (bytes[1] == 0x00) return "A";
                if (bytes[1] == 0x01) return "B";
            }

            return string.Empty;
        }

        static double MeasureSyntheticCid(byte[] bytes) {
            if (bytes.Length != 2 || bytes[0] != 0x01) {
                return 0;
            }

            return bytes[1] == 0x00 ? 600 : 400;
        }
    }

    [Fact]
    public void TextContentParser_RestoresTextStateAcrossGraphicsStateStack() {
        const string content = "BT /F1 10 Tf 0 0 Td (A) Tj ET q BT /F2 20 Tf 10 Tc 50 Tz 5 Ts 10 0 Td (B) Tj ET Q BT 20 0 Td (C) Tj ET";

        var spans = TextContentParser.Parse(
            content,
            static (_, bytes) => Encoding.ASCII.GetString(bytes),
            static (font, bytes) => font == "F2" ? (bytes?.Length ?? 0) * 1000 : (bytes?.Length ?? 0) * 500);

        PdfTextSpan first = Assert.Single(spans, span => span.Text == "A");
        PdfTextSpan scoped = Assert.Single(spans, span => span.Text == "B");
        PdfTextSpan restored = Assert.Single(spans, span => span.Text == "C");

        Assert.Equal("F1", first.FontResource);
        Assert.Equal(10, first.FontSize);
        Assert.Equal(5, first.Advance, 3);

        Assert.Equal("F2", scoped.FontResource);
        Assert.Equal(20, scoped.FontSize);
        Assert.Equal(5, scoped.Y, 3);

        Assert.Equal("F1", restored.FontResource);
        Assert.Equal(10, restored.FontSize);
        Assert.Equal(20, restored.X, 3);
        Assert.Equal(0, restored.Y, 3);
        Assert.Equal(5, restored.Advance, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_UsesMarkedContentActualText() {
        byte[] bytes = BuildSingleStreamPdf(
            "/Span << /ActualText <FEFF005A00650064> >> BDC\n" +
            "BT\n/F1 12 Tf\n72 720 Td\n(X) Tj\nET\n" +
            "EMC\n");

        var span = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetTextSpans());

        Assert.Equal("Zed", span.Text);
        Assert.Equal(72, span.X, 3);
        Assert.Equal(720, span.Y, 3);
        Assert.Equal(PdfWriter.EstimateSimpleTextWidth("X", PdfStandardFont.Helvetica, 12), span.Advance, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsLittleEndianMarkedContentActualText() {
        byte[] bytes = BuildSingleStreamPdf(
            "/Span << /ActualText <FFFE5A0065006400> >> BDC\n" +
            "BT\n/F1 12 Tf\n72 720 Td\n(X) Tj\nET\n" +
            "EMC\n");

        var span = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetTextSpans());

        Assert.Equal("Zed", span.Text);
        Assert.Equal(72, span.X, 3);
        Assert.Equal(720, span.Y, 3);
        Assert.Equal(PdfWriter.EstimateSimpleTextWidth("X", PdfStandardFont.Helvetica, 12), span.Advance, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsUtf8MarkedContentActualText() {
        byte[] bytes = BuildSingleStreamPdf(
            "/Span << /ActualText <EFBBBF5A6564> >> BDC\n" +
            "BT\n/F1 12 Tf\n72 720 Td\n(X) Tj\nET\n" +
            "EMC\n");

        var span = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetTextSpans());

        Assert.Equal("Zed", span.Text);
        Assert.Equal(72, span.X, 3);
        Assert.Equal(720, span.Y, 3);
        Assert.Equal(PdfWriter.EstimateSimpleTextWidth("X", PdfStandardFont.Helvetica, 12), span.Advance, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ResolvesMarkedContentActualTextFromPropertiesResource() {
        byte[] bytes = BuildSingleStreamPdfWithMarkedContentProperties(
            "/Span /MC0 BDC\n" +
            "BT\n/F1 12 Tf\n72 720 Td\n(X) Tj\nET\n" +
            "EMC\n");

        var span = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetTextSpans());

        Assert.Equal("Resource Zed", span.Text);
        Assert.Equal(72, span.X, 3);
        Assert.Equal(720, span.Y, 3);
        Assert.Equal(PdfWriter.EstimateSimpleTextWidth("X", PdfStandardFont.Helvetica, 12), span.Advance, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_HonorsEmptyMarkedContentActualText() {
        byte[] bytes = BuildSingleStreamPdf(
            "/Span << /ActualText <> >> BDC\n" +
            "BT\n/F1 12 Tf\n72 760 Td\n(Decorative glyph text) Tj\nET\n" +
            "EMC\n" +
            "BT\n/F1 12 Tf\n72 720 Td\n(Body text) Tj\nET\n");

        var page = PdfReadDocument.Open(bytes).Pages[0];
        var span = Assert.Single(page.GetTextSpans());

        Assert.Equal("Body text", span.Text);
        Assert.Equal(72, span.X, 3);
        Assert.Equal(720, span.Y, 3);
        Assert.DoesNotContain("Decorative glyph text", page.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_HonorsEmptyMarkedContentActualTextFromPropertiesResource() {
        byte[] bytes = BuildSingleStreamPdfWithMarkedContentProperties(
            "/Span /MC0 BDC\n" +
            "BT\n/F1 12 Tf\n72 760 Td\n(Decorative resource text) Tj\nET\n" +
            "EMC\n" +
            "BT\n/F1 12 Tf\n72 720 Td\n(Body text) Tj\nET\n",
            "<>");

        var page = PdfReadDocument.Open(bytes).Pages[0];
        var span = Assert.Single(page.GetTextSpans());

        Assert.Equal("Body text", span.Text);
        Assert.Equal(72, span.X, 3);
        Assert.Equal(720, span.Y, 3);
        Assert.DoesNotContain("Decorative resource text", page.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_IgnoresArtifactMarkedContent() {
        byte[] bytes = BuildSingleStreamPdf(
            "/Artifact BMC\n" +
            "BT\n/F1 12 Tf\n72 760 Td\n(Decorative header) Tj\nET\n" +
            "EMC\n" +
            "BT\n/F1 12 Tf\n72 720 Td\n(Body text) Tj\nET\n");

        var page = PdfReadDocument.Open(bytes).Pages[0];
        var span = Assert.Single(page.GetTextSpans());

        Assert.Equal("Body text", span.Text);
        Assert.Equal(72, span.X, 3);
        Assert.Equal(720, span.Y, 3);
        Assert.DoesNotContain("Decorative header", page.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsSimpleFontEncodingDifferences() {
        byte[] bytes = BuildPdfWithFontEncodingDifferences();

        var spans = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans();

        var span = Assert.Single(spans, item => item.Text == "Z \u20AC\u0104A");
        Assert.Equal(72, span.X, 3);
        Assert.Equal(720, span.Y, 3);
        Assert.Equal(PdfWriter.EstimateSimpleTextWidth("Z \u20AC\u0104A", PdfStandardFont.Helvetica, 12), span.Advance, 3);
    }

    [Theory]
    [InlineData("65 /uD800")]
    [InlineData("65 /uDFFF")]
    [InlineData("65 /uniD800")]
    public void PdfReadPage_GetTextSpans_RejectsSurrogateGlyphNames(string differences) {
        byte[] bytes = BuildPdfWithFontEncodingDifferences(differences, "41");

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetTextSpans());

        Assert.Equal("A", span.Text);
    }

    [Fact]
    public void TextContentParser_AdvancesRunsThroughActiveTextMatrix() {
        const string content = "BT /F1 10 Tf 2 0 0 2 10 20 Tm (A) Tj (B) Tj ET";

        var spans = TextContentParser.Parse(
            content,
            static (_, bytes) => Encoding.ASCII.GetString(bytes),
            static (_, bytes) => (bytes?.Length ?? 0) * 1000);

        PdfTextSpan first = Assert.Single(spans, span => span.Text == "A");
        PdfTextSpan second = Assert.Single(spans, span => span.Text == "B");
        Assert.Equal(10, first.X, 3);
        Assert.Equal(20, first.Y, 3);
        Assert.Equal(20, first.Advance, 3);
        Assert.Equal(30, second.X, 3);
        Assert.Equal(20, second.Y, 3);
    }

}

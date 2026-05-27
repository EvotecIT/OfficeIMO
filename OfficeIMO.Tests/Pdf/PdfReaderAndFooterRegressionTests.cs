using System;
using System.Reflection;
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

        string content = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/BaseFont /Helvetica-Bold", content);
    }

    [Fact]
    public void FooterFontResource_IsNotEmittedWhenFooterIsDisabled() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            FooterFont = PdfStandardFont.CourierBold,
            ShowPageNumbers = false
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("Body text only."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/BaseFont /Helvetica", content);
        Assert.DoesNotContain("/BaseFont /Courier-Bold", content);
        Assert.DoesNotContain("/F5", content);
    }

    [Fact]
    public void FooterSegments_RenderWithoutShowPageNumbersFlag() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.TimesRoman,
            FooterFont = PdfStandardFont.HelveticaBold,
            FooterAlign = PdfAlign.Left,
            FooterSegments = new System.Collections.Generic.List<FooterSegment> {
                new FooterSegment(FooterSegmentKind.Text, "Direct footer"),
                new FooterSegment(FooterSegmentKind.Text, " "),
                new FooterSegment(FooterSegmentKind.PageNumber)
            }
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("Body text only."))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        string pageText = new string(pdf.GetPage(1).Text.Where(c => !char.IsWhiteSpace(c)).ToArray());
        Assert.Contains("Directfooter1", pageText, StringComparison.OrdinalIgnoreCase);

        string content = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/BaseFont /Helvetica-Bold", content);
    }

    [Fact]
    public void FooterSegments_ValidateFooterPlacementWithoutShowPageNumbersFlag() {
        var options = new PdfOptions {
            MarginBottom = 20,
            FooterOffsetY = 21,
            FooterSegments = new System.Collections.Generic.List<FooterSegment> {
                new FooterSegment(FooterSegmentKind.Text, "Direct footer")
            }
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(options)
                .Paragraph(p => p.Text("Body text only."))
                .ToBytes());

        Assert.Contains("PDF footer offset must not exceed the bottom margin when footer content is enabled.", exception.Message, StringComparison.Ordinal);
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

        var span = Assert.Single(PdfReadDocument.Load(bytes).Pages[0].GetTextSpans());

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

        var span = Assert.Single(PdfReadDocument.Load(bytes).Pages[0].GetTextSpans());

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

        var span = Assert.Single(PdfReadDocument.Load(bytes).Pages[0].GetTextSpans());

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

        var page = PdfReadDocument.Load(bytes).Pages[0];
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

        var page = PdfReadDocument.Load(bytes).Pages[0];
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

        var page = PdfReadDocument.Load(bytes).Pages[0];
        var span = Assert.Single(page.GetTextSpans());

        Assert.Equal("Body text", span.Text);
        Assert.Equal(72, span.X, 3);
        Assert.Equal(720, span.Y, 3);
        Assert.DoesNotContain("Decorative header", page.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsSimpleFontEncodingDifferences() {
        byte[] bytes = BuildPdfWithFontEncodingDifferences();

        var spans = PdfReadDocument.Load(bytes).Pages[0].GetTextSpans();

        var span = Assert.Single(spans, item => item.Text == "Z \u20AC\u0104A");
        Assert.Equal(72, span.X, 3);
        Assert.Equal(720, span.Y, 3);
        Assert.Equal(PdfWriter.EstimateSimpleTextWidth("Z \u20AC\u0104A", PdfStandardFont.Helvetica, 12), span.Advance, 3);
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

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsPagesWithContentStreamArrays() {
        byte[] bytes = BuildPdfWithContentStreamArray();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello world", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_PreservesTextStateAcrossContentStreamArrays() {
        byte[] bytes = BuildPdfWithSplitTextStateContentStreamArray();

        var spans = PdfReadDocument.Load(bytes).Pages[0].GetTextSpans();

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

        var document = PdfReadDocument.Load(bytes);

        Assert.Equal(2, document.Pages.Count);
        Assert.NotEqual(document.Pages[0].ObjectNumber, document.Pages[1].ObjectNumber);
    }

    [Fact]
    public void PdfReadDocument_Load_PreservesPagesWithDistinctReferencedContentArrays() {
        byte[] bytes = BuildPdfWithDistinctReferencedContentArrays();

        var document = PdfReadDocument.Load(bytes);

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
    public void PdfTextExtractor_InternalExtractor_IgnoresDirectFormCycles() {
        var state = BuildDirectFormCycleState();
        var method = typeof(PdfTextExtractor).GetMethod("ExtractTextFromContentStream", BindingFlags.NonPublic | BindingFlags.Static);

        string text = Assert.IsType<string>(method!.Invoke(null, new object?[] {
            "/Fx Do",
            state.Resources,
            state.Objects,
            new Dictionary<int, string>(),
            new HashSet<PdfStream>()
        }));

        Assert.Equal(string.Empty, text);
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

        string text = PdfReadDocument.Load(bytes).ExtractText();

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

        var page = PdfReadDocument.Load(bytes).Pages[0].ExtractStructured(new PdfTextLayoutOptions {
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
    public void PdfReadPage_GetTextSpans_ReadsEscapedLiteralStrings() {
        byte[] bytes = BuildSingleStreamPdf("BT\n/F1 12 Tf\n72 720 Td\n(Hello\\040octal\\041) Tj\n( Line\\nbreak) Tj\nET\n");

        var spans = PdfReadDocument.Load(bytes).Pages[0].GetTextSpans();

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

        var document = PdfReadDocument.Load(bytes);

        Assert.Equal("Hello meta!", document.Metadata.Title);
        Assert.Equal("OfficeIMOTeam", document.Metadata.Author);
    }

    [Fact]
    public void PdfReadDocument_Metadata_ReadsLiteralStringsContainingStreamSubstrings() {
        byte[] bytes = BuildPdfWithLiteralMetadata("(mainstream title)", "(upstream author)");

        var document = PdfReadDocument.Load(bytes);

        Assert.Equal("mainstream title", document.Metadata.Title);
        Assert.Equal("upstream author", document.Metadata.Author);
    }

    [Fact]
    public void PdfReadDocument_Metadata_ReadsLiteralStringsContainingStandaloneParserKeywords() {
        byte[] bytes = BuildPdfWithLiteralMetadata("(stream title)", "(endobj author)");

        var document = PdfReadDocument.Load(bytes);

        Assert.Equal("stream title", document.Metadata.Title);
        Assert.Equal("endobj author", document.Metadata.Author);
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
    public void PdfTextExtractor_ExtractAllText_ReadsLzwEncodedContentStreams() {
        byte[] bytes = BuildPdfWithLzwEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello LZW stream) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello LZW stream", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsLzwEncodedContentStreams() {
        byte[] bytes = BuildPdfWithLzwEncodedStream("BT\n/F1 12 Tf\n72 720 Td\n(Hello LZW spans) Tj\nET\n");

        var doc = PdfReadDocument.Load(bytes);

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
    public void PdfReadPage_GetTextSpans_AppliesScaledFormTransformsInOrder() {
        byte[] bytes = BuildPdfWithScaledFormMatrix();

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Scaled form");
        Assert.Equal(26, span.X, 3);
        Assert.Equal(42, span.Y, 3);
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

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsStreamsContainingEndobjLiterals() {
        byte[] bytes = BuildSingleStreamPdf("BT\n/F1 12 Tf\n72 720 Td\n(Hello endobj marker) Tj\nET\n");

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Hello endobj marker", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsStreamsContainingEndobjLiterals() {
        byte[] bytes = BuildSingleStreamPdf("BT\n/F1 12 Tf\n72 720 Td\n(Hello endobj marker) Tj\nET\n");

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans());
        Assert.Equal("Hello endobj marker", span.Text);
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
        const string streamOne = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\nET";
        const string streamTwo = "\nBT\n/F1 12 Tf\n72 720 Td\n( world) Tj\nET";
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

    private static byte[] BuildPdfWithSplitTextStateContentStreamArray() {
        const string streamOne = "BT\n/F1 12 Tf\n72 720 Td";
        const string streamTwo = "\n(Split state) Tj\nET";
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

    private static byte[] BuildPdfWithCyclicKidsReferences() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello cyclic kids) Tj\nET\n";
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
            "<< /Type /Pages /Parent 2 0 R /Kids [2 0 R 4 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 3 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 6 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "6 0 obj",
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

    private static byte[] BuildPdfWithIndirectContentArrayObject() {
        const string streamOne = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\nET";
        const string streamTwo = "\nBT\n/F1 12 Tf\n72 720 Td\n( world) Tj\nET";
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

    private static void AssertSecondLineStartsAtFirstLineX(byte[] bytes) {
        var spans = PdfReadDocument.Load(bytes).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan first = Assert.Single(spans, span => span.Text == "First");
        PdfTextSpan second = Assert.Single(spans, span => span.Text == "Second");

        Assert.Equal(first.X, second.X, 2);
        Assert.True(second.Y < first.Y, $"Expected second line Y {second.Y} to be below first line Y {first.Y}.");
    }

    private static byte[] BuildPdfWithSingleQuoteOperator() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\n( world) '\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithDoubleQuoteOperator() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\n0 0 ( world) \"\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithDoubleQuoteLineAdvanceOperator() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(First) Tj\n0 0 (Second) \"\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithTDTextPositioning() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(First) Tj\n0 -14 TD\n(Second) Tj\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithInitialTdTextPositioning() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(First) Tj\nET\nBT\n/F1 12 Tf\n110 720 Td\n(Second) Tj\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithRepeatedTDTextPositioning() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(First) Tj\n0 -14 TD\n0 -14 TD\n(Second) Tj\nET\n";
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

    private static byte[] BuildPdfWithLzwEncodedStream(string streamContent) {
        byte[] encodedBytes = EncodeLzw(Encoding.ASCII.GetBytes(streamContent));
        return BuildSingleStreamPdf(encodedBytes, "/Filter /LZWDecode");
    }

    private static byte[] BuildPdfWithAscii85AndLzwEncodedStream(string streamContent, int earlyChange) {
        byte[] lzwBytes = EncodeLzw(Encoding.ASCII.GetBytes(streamContent), earlyChange);
        byte[] encodedBytes = Encoding.ASCII.GetBytes(EncodeAscii85(lzwBytes));
        return BuildSingleStreamPdf(encodedBytes, $"/Filter [/A85 /LZW] /DecodeParms [null << /EarlyChange {earlyChange} >>]");
    }

    private static byte[] BuildPdfWithLzwPredictorEncodedStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] predictedBytes = EncodeUpPredictedRows(streamBytes);
        byte[] encodedBytes = EncodeLzw(predictedBytes);
        return BuildSingleStreamPdf(encodedBytes, $"/Filter /LZWDecode /DecodeParms << /Predictor 12 /Columns {streamBytes.Length} >>");
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

    private static byte[] BuildSingleStreamPdfWithMarkedContentProperties(string streamContent, string actualTextPdfString = "<FEFF005200650073006F00750072006300650020005A00650064>") {
        streamContent = streamContent.TrimEnd('\n');
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
            $"<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> /Properties << /MC0 << /ActualText {actualTextPdfString} >> >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithFontEncodingDifferences() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n<4142434445> Tj\nET\n";
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
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /BaseEncoding /WinAnsiEncoding /Differences [65 /Z /space /Euro /uni0104 /A.alt] >> >>",
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

    private static byte[] EncodeLzw(byte[] input, int earlyChange = 1) {
        earlyChange = earlyChange == 0 ? 0 : 1;
        var dictionary = new Dictionary<string, int>();
        for (int i = 0; i < 256; i++) {
            dictionary[Convert.ToBase64String(new[] { (byte)i })] = i;
        }

        using var output = new MemoryStream();
        var writer = new LzwBitWriter(output);
        int nextCode = 258;
        int codeSize = 9;
        writer.WriteBits(256, codeSize);
        var current = new List<byte>();

        foreach (byte value in input) {
            var candidate = new List<byte>(current) { value };
            string candidateKey = Convert.ToBase64String(candidate.ToArray());
            if (dictionary.ContainsKey(candidateKey)) {
                current = candidate;
                continue;
            }

            writer.WriteBits(dictionary[Convert.ToBase64String(current.ToArray())], codeSize);
            if (nextCode <= 4095) {
                dictionary[candidateKey] = nextCode++;
                if (codeSize < 12 && nextCode + earlyChange >= (1 << codeSize)) {
                    codeSize++;
                }
            }

            current = new List<byte> { value };
        }

        if (current.Count > 0) {
            writer.WriteBits(dictionary[Convert.ToBase64String(current.ToArray())], codeSize);
        }

        writer.WriteBits(257, codeSize);
        writer.Flush();
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

    private sealed class LzwBitWriter {
        private readonly Stream _stream;
        private int _currentByte;
        private int _bitCount;

        public LzwBitWriter(Stream stream) {
            _stream = stream;
        }

        public void WriteBits(int value, int bitCount) {
            for (int i = bitCount - 1; i >= 0; i--) {
                _currentByte = (_currentByte << 1) | ((value >> i) & 1);
                _bitCount++;
                if (_bitCount == 8) {
                    _stream.WriteByte((byte)_currentByte);
                    _currentByte = 0;
                    _bitCount = 0;
                }
            }
        }

        public void Flush() {
            if (_bitCount == 0) {
                return;
            }

            _stream.WriteByte((byte)(_currentByte << (8 - _bitCount)));
            _currentByte = 0;
            _bitCount = 0;
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

    private static byte[] BuildPdfWithScaledFormMatrix() {
        const string pageContent = "q\n2 0 0 2 10 20 cm\n/Fx Do\nQ\n";
        const string formContent = "BT\n/F1 12 Tf\n3 4 Td\n(Scaled form) Tj\nET\n";
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
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Matrix [1 0 0 1 5 7] /Resources 9 0 R /Length {formStreamLength} >>",
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

    private static byte[] BuildPdfWithTwoDirectContentPages() {
        const string pageOne = "BT\n/F1 12 Tf\n72 720 Td\n(Direct page one) Tj\nET\n";
        const string pageTwo = "BT\n/F1 12 Tf\n72 720 Td\n(Direct page two) Tj\nET\n";

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 612 792] /Resources 5 0 R >>",
            "endobj",
            "3 0 obj",
            $"<< /Type /Page /Parent 2 0 R /Contents << /Length {Encoding.ASCII.GetByteCount(pageOne)} >>\nstream\n{pageOne.TrimEnd('\n')}\nendstream >>",
            "endobj",
            "4 0 obj",
            $"<< /Type /Page /Parent 2 0 R /Contents << /Length {Encoding.ASCII.GetByteCount(pageTwo)} >>\nstream\n{pageTwo.TrimEnd('\n')}\nendstream >>",
            "endobj",
            "5 0 obj",
            "<< /Font 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /F1 7 0 R >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithDistinctReferencedContentArrays() {
        const string sharedStream = "BT\n/F1 12 Tf\n72 720 Td\n(Shared stream page) Tj\nET\n";
        int sharedLength = Encoding.ASCII.GetByteCount(sharedStream);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 612 792] /Resources 8 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>",
            "endobj",
            "5 0 obj",
            "[7 0 R]",
            "endobj",
            "6 0 obj",
            "[7 0 R]",
            "endobj",
            "7 0 obj",
            $"<< /Length {sharedLength} >>",
            "stream",
            sharedStream.TrimEnd('\n'),
            "endstream",
            "endobj",
            "8 0 obj",
            "<< /Font 9 0 R >>",
            "endobj",
            "9 0 obj",
            "<< /F1 10 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
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

    private static PdfReadPage CreatePdfReadPageWithDirectFormCycle() {
        var state = BuildDirectFormCycleState();
        var pageDict = new PdfDictionary();
        pageDict.Items["Type"] = new PdfName("Page");
        pageDict.Items["Resources"] = state.Resources;

        var contents = new PdfArray();
        contents.Items.Add(new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes("/Fx Do")));
        pageDict.Items["Contents"] = contents;

        var ctor = typeof(PdfReadPage).GetConstructor(
            BindingFlags.Instance | BindingFlags.NonPublic,
            binder: null,
            new[] { typeof(int), typeof(PdfDictionary), typeof(Dictionary<int, PdfIndirectObject>) },
            modifiers: null);

        return Assert.IsType<PdfReadPage>(ctor!.Invoke(new object[] { 1, pageDict, state.Objects }));
    }

    private static (PdfDictionary Resources, Dictionary<int, PdfIndirectObject> Objects) BuildDirectFormCycleState() {
        var xObjects = new PdfDictionary();
        var resources = new PdfDictionary();
        resources.Items["XObject"] = xObjects;

        var formDict = new PdfDictionary();
        formDict.Items["Subtype"] = new PdfName("Form");
        formDict.Items["Resources"] = resources;

        var formStream = new PdfStream(formDict, Encoding.ASCII.GetBytes("/Fx Do"));
        xObjects.Items["Fx"] = formStream;

        return (resources, new Dictionary<int, PdfIndirectObject>());
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

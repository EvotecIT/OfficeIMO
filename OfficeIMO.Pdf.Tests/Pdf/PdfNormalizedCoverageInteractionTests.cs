using System;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNormalizedCoverageInteractionTests {
    [Theory]
    [InlineData(false, "\n")]
    [InlineData(true, "\n")]
    [InlineData(false, "\r")]
    [InlineData(true, "\r")]
    [InlineData(false, "\t")]
    [InlineData(true, "\t")]
    public void NormalizedPrimaryCoverageCoexistsWithAuthoredLayoutSeparators(bool named, string separator) {
        string source = "A\u0301" + separator + "A\u0301";
        (PdfDocument document, PdfConversionReport report) = CreateDocument(source, named, unrelatedFallback: false);

        Assert.Empty(document.AnalyzeTextEncoding());
        byte[] bytes = document.ToBytes();
        Assert.NotEmpty(bytes);
        // CR/LF/tab are PDF layout separators, so extraction may canonicalize
        // their whitespace. Every original NFD source cluster remains once.
        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Equal("A\u0301A\u0301", new string(extracted.Where(character => !char.IsWhiteSpace(character)).ToArray()));
        report.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnrelatedFallbackCannotInvalidateCompletePrimaryNormalization(bool named) {
        const string source = "A\u0301";
        (PdfDocument document, PdfConversionReport report) = CreateDocument(source, named, unrelatedFallback: true);

        Assert.Empty(document.AnalyzeTextEncoding());
        byte[] bytes = document.ToBytes();
        Assert.NotEmpty(bytes);
        Assert.Equal(source, PdfReadDocument.Open(bytes).ExtractText().Trim());
        report.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NeededFallbackPreservesTheCompleteNormalizedPrimaryCluster(bool named) {
        const string source = "A\u0301B";
        (PdfDocument document, PdfConversionReport report) = CreateDocument(source, named, unrelatedFallback: true);

        Assert.Empty(document.AnalyzeTextEncoding());
        byte[] bytes = document.ToBytes();
        Assert.NotEmpty(bytes);
        // The primary must retain the original A+acute cluster while B uses
        // the real fallback; a whole-primary bypass cannot satisfy this case.
        Assert.Equal(source, PdfReadDocument.Open(bytes).ExtractText().Trim());
        report.RequireNoLoss();
    }

    [Theory]
    [InlineData("A\u0301\r\nA\u0001", 5, "U+0001", true)]
    [InlineData("A\u0301\tZ", 3, "U+005A", false)]
    public void LayoutSegmentDiagnosticsRetainOriginalSourceOffsetsAndStrictFailures(
        string source, int index, string codePoint, bool control) {
        var (document, _) = CreateDocument(source, named: false, unrelatedFallback: false);

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(document.AnalyzeTextEncoding());
        Assert.Equal(index, diagnostic.Index);
        Assert.Equal(codePoint, diagnostic.CodePoint);
        Assert.Equal(control, diagnostic.IsControlCharacter);
        Assert.Equal(control ? "unsupported-control-character" : "unsupported-text-glyph", diagnostic.Code);
        ArgumentException error = Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());
        Assert.Equal(index, error.Data["index"]);
        Assert.Equal(codePoint, error.Data["codePoint"]);
    }

    private static (PdfDocument Document, PdfConversionReport Report) CreateDocument(
        string source, bool named, bool unrelatedFallback) {
        const string family = "Normalized interaction";
        // Same accepted no-GSUB fixture as PdfNormalizedEncodingCoverageTests:
        // A and NFC A-acute exist, but the original combining acute does not.
        byte[] primary = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true, 'A', 0x00C1);
        var report = new PdfConversionReport();
        PdfCreateOptions options = (named
            ? new PdfCreateOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family, primary))
            : new PdfCreateOptions().EmbedStandardFont(PdfStandardFont.Helvetica, primary, family))
            .SetTextShapingMode(PdfTextShapingMode.OpenTypeLigatures)
            .ReportDiagnosticsTo(report);
        if (unrelatedFallback) {
            byte[] secondary = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true, 'B');
            var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("Unrelated B fallback", secondary)
            });
            // The candidate is a valid font for B, not a replacement for A-acute.
            Assert.True(fallbacks.PlanText("B").IsFullyCovered);
            options.RegisterEmbeddedFontFallbacks(fallbacks);
        }
        PdfDocument document = PdfDocument.Create(options).Paragraph(paragraph => {
            if (named) paragraph.FontFamily(family);
            paragraph.Text(source);
        });
        return (document, report);
    }
}

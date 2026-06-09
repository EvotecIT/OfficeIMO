using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class RichParagraphWrappingTests {
        [Fact]
        public void WinAnsiEncoding_EncodesSupportedWindows1252CharactersWithoutFallback() {
            var bytes = PdfWinAnsiEncoding.Encode("\u20AC\u2022\u201C\u201D\u0152\u0178");

            Assert.Equal(new byte[] { 0x80, 0x95, 0x93, 0x94, 0x8C, 0x9F }, bytes);
            Assert.True(PdfWinAnsiEncoding.CanEncode("Invoice \u20AC \u2022", out int unsupportedIndex));
            Assert.Equal(-1, unsupportedIndex);
        }

        [Fact]
        public void EstimateSimpleTextWidth_UsesWinAnsiPunctuationWidths() {
            double width = InvokePrivateFontMethod<double>(
                "EstimateSimpleTextWidth",
                "\u201CWait\u201D\u2014ok\u2026",
                PdfStandardFont.TimesRoman,
                10.0);

            Assert.Equal(58.32, width, 2);
        }

        [Fact]
        public void EstimateSimpleTextWidth_UsesWinAnsiAccentedLetterWidths() {
            double width = InvokePrivateFontMethod<double>(
                "EstimateSimpleTextWidth",
                "r\u00E9sum\u00E9",
                PdfStandardFont.TimesRoman,
                10.0);

            Assert.Equal(28.88, width, 2);
        }

        [Fact]
        public void WinAnsiEncoding_RejectsUnsupportedGeneratedTextWithClearDiagnostic() {
            var exception = Assert.ThrowsAny<ArgumentException>(() =>
                PdfDocument.Create()
                    .Paragraph(p => p.Text("Snowman \u2603"))
                    .ToBytes());

            Assert.Contains("U+2603", exception.Message, StringComparison.Ordinal);
            Assert.Contains("PDF WinAnsiEncoding", exception.Message, StringComparison.Ordinal);
            Assert.Contains("Embedded Unicode fonts", exception.Message, StringComparison.Ordinal);
            Assert.Equal("unsupported-text-glyph", exception.Data["code"]);
            Assert.Equal("PdfParagraph", exception.Data["source"]);
            Assert.Equal("PdfParagraph[0].Run[0]", exception.Data["location"]);
            Assert.Equal(0, exception.Data["runIndex"]);
            Assert.Equal("PDF WinAnsiEncoding", exception.Data["encoding"]);
            Assert.Equal("Embedded Unicode fonts are required for this text.", exception.Data["remediation"]);
            Assert.Equal(1, exception.Data["diagnosticsCount"]);
            Assert.False(PdfWinAnsiEncoding.CanEncode("Snowman \u2603", out int unsupportedIndex));
            Assert.Equal(8, unsupportedIndex);
        }

        [Fact]
        public void WinAnsiEncoding_RejectsControlCharactersWithLayoutDiagnostic() {
            var exception = Assert.Throws<ArgumentException>(() => PdfWinAnsiEncoding.Encode("Alpha\tBeta"));

            Assert.Contains("control character U+0009", exception.Message, StringComparison.Ordinal);
            Assert.Contains("layout", exception.Message, StringComparison.Ordinal);
            Assert.False(PdfWinAnsiEncoding.CanEncode("Alpha\tBeta", out int unsupportedIndex));
            Assert.Equal(5, unsupportedIndex);
        }

        [Fact]
        public void TextDiagnostics_ReportsUnsupportedGlyphAsSharedConversionWarning() {
            var diagnostic = Assert.Single(PdfTextDiagnostics.AnalyzeWinAnsiText("Snowman \u2603", "paragraph"));

            Assert.Equal("paragraph", diagnostic.Source);
            Assert.Equal(8, diagnostic.Index);
            Assert.Equal("U+2603", diagnostic.CodePoint);
            Assert.Equal("\u2603", diagnostic.Text);
            Assert.False(diagnostic.IsControlCharacter);
            Assert.Equal("PDF WinAnsiEncoding", diagnostic.Encoding);
            Assert.Equal("Embedded Unicode fonts are required for this text.", diagnostic.Remediation);
            Assert.Equal("unsupported-text-glyph", diagnostic.Code);
            Assert.Contains("Embedded Unicode fonts", diagnostic.Message, StringComparison.Ordinal);

            PdfConversionWarning warning = diagnostic.ToConversionWarning("OfficeIMO.Tests");

            Assert.Equal("OfficeIMO.Tests", warning.Converter);
            Assert.Equal("unsupported-text-glyph", warning.Code);
            Assert.Equal(PdfConversionWarningSeverity.Error, warning.Severity);
            Assert.Equal(PdfLayoutDiagnosticKind.SimplifiedContent, warning.LayoutDiagnostic!.Kind);
            Assert.Equal("U+2603", warning.Details["codePoint"]);
            Assert.Equal("8", warning.Details["index"]);
            Assert.Equal("PDF WinAnsiEncoding", warning.Details["encoding"]);
            Assert.Equal("Embedded Unicode fonts are required for this text.", warning.Details["remediation"]);
        }

        [Fact]
        public void TextDiagnostics_ReportsControlCharactersButIgnoresLayoutControls() {
            IReadOnlyList<PdfTextEncodingDiagnostic> literalTabDiagnostics = PdfTextDiagnostics.AnalyzeWinAnsiText("Alpha\tBeta", "literal");

            Assert.Empty(literalTabDiagnostics);

            var controlDiagnostic = Assert.Single(PdfTextDiagnostics.AnalyzeWinAnsiText("Alpha\u0001Beta", "literal"));

            Assert.True(controlDiagnostic.IsControlCharacter);
            Assert.Equal("unsupported-control-character", controlDiagnostic.Code);
            Assert.Equal("U+0001", controlDiagnostic.CodePoint);
            Assert.Equal(string.Empty, controlDiagnostic.Text);
            Assert.Equal("PDF text output", controlDiagnostic.Encoding);
            Assert.Equal("Use paragraphs, line breaks, tables, or spacing primitives for layout instead of literal control characters.", controlDiagnostic.Remediation);

            IReadOnlyList<PdfTextEncodingDiagnostic> runDiagnostics = PdfTextDiagnostics.AnalyzeWinAnsiTextRuns(
                new[] {
                    TextRun.Normal("Alpha"),
                    TextRun.Tab(PdfTabLeaderStyle.Dots),
                    TextRun.LineBreak(),
                    TextRun.Normal("Beta")
                },
                "runs");

            Assert.Empty(runDiagnostics);
        }

        [Fact]
        public void TextDiagnostics_ReportsRunLocationForRichTextDiagnostics() {
            IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeGeneratedTextRuns(
                new[] {
                    TextRun.Normal("Alpha"),
                    TextRun.Tab(PdfTabLeaderStyle.Dots),
                    TextRun.LineBreak(),
                    TextRun.Normal("Beta \u2603")
                },
                new PdfOptions(),
                PdfStandardFont.Helvetica,
                "paragraph",
                "PdfParagraph[0]");

            PdfTextEncodingDiagnostic diagnostic = Assert.Single(diagnostics);

            Assert.Equal("paragraph", diagnostic.Source);
            Assert.Equal("PdfParagraph[0].Run[3]", diagnostic.Location);
            Assert.Equal(3, diagnostic.RunIndex);
            Assert.Equal("U+2603", diagnostic.CodePoint);
            Assert.Equal("PDF WinAnsiEncoding", diagnostic.Encoding);

            PdfConversionWarning warning = diagnostic.ToConversionWarning("OfficeIMO.Tests");

            Assert.Equal("3", warning.Details["runIndex"]);
            Assert.Equal("PDF WinAnsiEncoding", warning.Details["encoding"]);
        }

        [Fact]
        public void TextDiagnostics_AllowsSupportedWindows1252Text() {
            IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeWinAnsiText("Invoice \u20AC \u2022 r\u00E9sum\u00E9", "paragraph");

            Assert.Empty(diagnostics);
        }

        [Fact]
        public void TextRun_TabLeaderRequiresExplicitTabRun() {
            var invalidLeaderException = Assert.Throws<ArgumentException>(() =>
                new TextRun("Alpha", tabLeader: PdfTabLeaderStyle.Dots));

            Assert.Contains("Tab leaders and alignment can only be applied to explicit tab runs.", invalidLeaderException.Message, StringComparison.Ordinal);

            var invalidEnumException = Assert.Throws<ArgumentException>(() =>
                TextRun.Tab((PdfTabLeaderStyle)99));

            Assert.Contains("PDF tab leader style must be None, Dots, Hyphens, or Underscores.", invalidEnumException.Message, StringComparison.Ordinal);

            var invalidAlignmentException = Assert.Throws<ArgumentException>(() =>
                TextRun.Tab(alignment: (PdfTabAlignment)99));

            Assert.Contains("PDF tab alignment must be Left, Center, Right, or DecimalSeparator.", invalidAlignmentException.Message, StringComparison.Ordinal);
        }
    }
}

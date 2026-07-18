using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class RichParagraphWrappingTests {
        [Fact]
        public void PdfOptions_ClonePreservesTextLineBreakCallback() {
            PdfTextLineBreakCallback callback = token => token == "alphaomega"
                ? new[] { 5 }
                : Array.Empty<int>();

            PdfOptions options = new PdfOptions().SetTextLineBreaks(callback);

            PdfOptions clone = options.Clone();

            Assert.Same(callback, clone.TextLineBreakCallback);
        }

        [Fact]
        public void WrapSimpleText_UsesTextLineBreakCallbackWithoutAddingHyphen() {
            const string text = "alphaomega";
            double maxWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "alpha", PdfStandardFont.Helvetica, 10.0) + 0.5D;
            var options = new PdfOptions().SetTextLineBreaks(token => token == text ? new[] { 5 } : Array.Empty<int>());

            var lines = InvokePrivateFontMethod<List<string>>("WrapSimpleTextForOptions", text, maxWidth, PdfStandardFont.Helvetica, 10.0, options);

            Assert.Equal(new[] { "alpha", "omega" }, lines);
            Assert.DoesNotContain("-", string.Concat(lines), StringComparison.Ordinal);
        }

        [Fact]
        public void WrapRichRuns_UsesTextLineBreakCallbackWithoutAddingHyphen() {
            const string text = "alphaomega";
            double maxWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "alpha", PdfStandardFont.Helvetica, 10.0) + 0.5D;
            var options = new PdfOptions().SetTextLineBreaks(token => token == text ? new[] { 5 } : Array.Empty<int>());

            object result = InvokePrivateFontMethod<object>(
                "WrapRichRunsCore",
                new[] { new TextRun(text) },
                maxWidth,
                10.0,
                PdfStandardFont.Helvetica,
                14.0,
                null!,
                36.0,
                options,
                null!);
            string[] lineTexts = ExtractLines(result)
                .Select(line => string.Concat(line.Select(ExtractText)))
                .ToArray();

            Assert.Equal(new[] { "alpha", "omega" }, lineTexts);
            Assert.DoesNotContain("-", string.Concat(lineTexts), StringComparison.Ordinal);
        }

        [Fact]
        public void GeneratedText_UsesTextLineBreakCallbackForLongUnspacedTokens() {
            int callbackCalls = 0;
            PdfTextLineBreakCallback callback = token => {
                callbackCalls++;
                return token == "alphaomega"
                    ? new[] { 5 }
                    : Array.Empty<int>();
            };

            byte[] bytes = PdfDocument.Create(new PdfOptions {
                    PageWidth = 80,
                    PageHeight = 180,
                    MarginLeft = 18,
                    MarginRight = 18,
                    MarginTop = 24,
                    MarginBottom = 24,
                    DefaultFontSize = 10,
                    CompressContentStreams = false
                })
                .TextLineBreaks(callback)
                .Paragraph(paragraph => paragraph.Text("alphaomega"))
                .ToBytes();

            string extracted = PdfReadDocument.Open(bytes).ExtractText();
            string[] lines = extracted
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            Assert.True(callbackCalls > 0);
            Assert.Contains("alpha", lines);
            Assert.Contains("omega", lines);
            Assert.DoesNotContain("alpha-", extracted, StringComparison.Ordinal);
        }

        [Fact]
        public void TextLineBreakCallback_SuppressesScriptSpecificLineBreakingWarningWhenFontCoversText() {
            string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
            if (fontPath == null) {
                return;
            }

            const string text = "\u0E20\u0E32\u0E29\u0E32\u0E44\u0E17\u0E22";
            byte[] fontData = File.ReadAllBytes(fontPath);
            PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMO Thai Test Font");
            if (PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram).Count > 0) {
                return;
            }

            var report = new PdfConversionReport();
            var options = new PdfOptions {
                    CompressContentStreams = false
                }
                .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
                .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Thai Test Font")
                .SetTextLineBreaks(token => token == text ? new[] { 4 } : Array.Empty<int>());

            byte[] bytes = PdfDocument.Create(options)
                .Paragraph(paragraph => paragraph.Text(text))
                .ToBytes();

            string extracted = PdfReadDocument.Open(bytes).ExtractText();

            Assert.Contains(text, extracted, StringComparison.Ordinal);
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-script-specific-line-breaking");
        }
    }
}

using System;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class RichParagraphWrappingHyphenationTests {
    [Fact]
    public void HyphenationDictionary_NormalizesEntriesAndHonorsPrefixSuffixLimits() {
        var dictionary = new PdfHyphenationLexicon(new[] {
            "ty-pog-ra-phy",
            "TYPOG-RAPHY"
        }, minimumPrefixLength: 2, minimumSuffixLength: 2);

        Assert.Equal(1, dictionary.Count);
        Assert.True(dictionary.Contains("Typography"));
        Assert.Equal(new[] { 2, 5, 7 }, dictionary.GetBreakpoints("typography"));
        Assert.Throws<ArgumentException>(() => new PdfHyphenationLexicon(new[] { "-invalid" }));
        Assert.Throws<ArgumentException>(() => new PdfHyphenationLexicon(new[] { "a-b" }, minimumPrefixLength: 2));
    }

    [Fact]
    public void GeneratedText_UsesFirstPartyHyphenationDictionary() {
        var dictionary = new PdfHyphenationLexicon(new[] { "typog-ra-phy-mile-stone" });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 118,
                PageHeight = 180,
                MarginLeft = 18,
                MarginRight = 18,
                MarginTop = 24,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .TextHyphenationDictionary(dictionary)
            .Paragraph(paragraph => paragraph.Text("typographymilestone"))
            .ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("typographymile-", extracted, StringComparison.Ordinal);
        Assert.Contains("stone", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfOptions_ClonePreservesTextHyphenationCallback() {
        PdfTextHyphenationCallback callback = token => token == "typographymilestone"
            ? new[] { 10 }
            : Array.Empty<int>();

        PdfOptions options = new PdfOptions().SetTextHyphenation(callback);

        PdfOptions clone = options.Clone();

        Assert.Same(callback, clone.TextHyphenationCallback);
    }

    [Fact]
    public void GeneratedText_UsesTextHyphenationCallbackForLongUnspacedTokens() {
        int callbackCalls = 0;
        PdfTextHyphenationCallback callback = token => {
            callbackCalls++;
            return token == "typographymilestone"
                ? new[] { 10 }
                : Array.Empty<int>();
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 118,
                PageHeight = 180,
                MarginLeft = 18,
                MarginRight = 18,
                MarginTop = 24,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .TextHyphenation(callback)
            .Paragraph(paragraph => paragraph.Text("typographymilestone"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.True(callbackCalls > 0);
        Assert.Contains("7479706F6772617068792D", raw, StringComparison.Ordinal);
        Assert.Contains("6D696C6573746F6E65", raw, StringComparison.Ordinal);
        Assert.Contains("typography-", extracted, StringComparison.Ordinal);
        Assert.Contains("milestone", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void GeneratedText_RollsBackPartialHyphenationWhenLaterChunkCannotFit() {
        PdfTextHyphenationCallback callback = token => token == "typographymilestone"
            ? new[] { 1 }
            : Array.Empty<int>();

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 118,
                PageHeight = 180,
                MarginLeft = 18,
                MarginRight = 18,
                MarginTop = 24,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .TextHyphenation(callback)
            .Paragraph(paragraph => paragraph.Text("typographymilestone"))
            .ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.DoesNotContain("t-", extracted, StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(extracted, "typography"));
    }

    [Fact]
    public void GeneratedText_BreaksLongDelimitedTokensAtRealDelimitersBeforeCharacterChunks() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 118,
                PageHeight = 180,
                MarginLeft = 18,
                MarginRight = 18,
                MarginTop = 24,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text("CASE-REVIEW-LINE-001"))
            .ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("CASE-REVIEW-", extracted, StringComparison.Ordinal);
        Assert.Contains("LINE-001", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain("CASE-REVIEW-LI", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void GeneratedText_PreservesDelimiterChunksWhenLaterSegmentNeedsCharacterWrapping() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 72,
                PageHeight = 180,
                MarginLeft = 18,
                MarginRight = 18,
                MarginTop = 24,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text("INV-DNS-LOCATOR-001"))
            .ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        string[] lines = extracted
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

        Assert.Contains("INV-", lines);
        Assert.Contains("DNS-", lines);
        Assert.Contains("LOCA", lines);
        Assert.Contains("TOR-", lines);
        Assert.Contains("001", lines);
        Assert.DoesNotContain("INV-D", lines);
        Assert.DoesNotContain("NS-LO", lines);
    }

    [Fact]
    public void GeneratedText_UsesAvailableWidthAfterIdentifierDelimiters() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 160,
                PageHeight = 180,
                MarginLeft = 18,
                MarginRight = 18,
                MarginTop = 24,
                MarginBottom = 24,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text(@"ACME\Service_2026AlphaIdentifier"))
            .ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        string[] lines = extracted
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

        Assert.Contains(lines, line => line.StartsWith(@"ACME\Service_2026", StringComparison.Ordinal));
        Assert.DoesNotContain(@"ACME\Service_", lines);
    }

    [Fact]
    public void GeneratedText_CarriesDelimiterToNextLineBeforeSplittingWordSegment() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 86,
                PageHeight = 180,
                MarginLeft = 18,
                MarginRight = 18,
                MarginTop = 24,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text("CASE-RECORD-STATIC-001"))
            .ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        string[] lines = extracted
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

        Assert.Contains("CASE-", lines);
        Assert.Contains("RECORD", lines);
        Assert.Contains("-STATIC-", lines);
        Assert.Contains("001", lines);
        Assert.DoesNotContain("RECOR", lines);
        Assert.DoesNotContain("D-", lines);
    }

    [Fact]
    public void GeneratedTableText_CarriesDelimitedWordSegmentBeforeCharacterFallback() {
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new System.Collections.Generic.List<double?> { 42, 70 };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 160,
                PageHeight = 180,
                MarginLeft = 18,
                MarginRight = 18,
                MarginTop = 24,
                MarginBottom = 24,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 9,
                CompressContentStreams = false
            })
            .Table(new[] {
                new[] { "Code", "Summary" },
                new[] { "DNS-RECORD-STATIC-001", "Contract clause inventory" }
            }, style: style)
            .ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        string[] lines = extracted
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

        Assert.Contains("RECORD", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain(lines, line => string.Equals(line, "RECOR", StringComparison.Ordinal));
        Assert.DoesNotContain(lines, line => string.Equals(line, "D-", StringComparison.Ordinal));
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }
}

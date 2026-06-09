using System;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class RichParagraphWrappingHyphenationTests {
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
        string extracted = PdfReadDocument.Load(bytes).ExtractText();

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

        string extracted = PdfReadDocument.Load(bytes).ExtractText();

        Assert.DoesNotContain("t-", extracted, StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(extracted, "typography"));
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

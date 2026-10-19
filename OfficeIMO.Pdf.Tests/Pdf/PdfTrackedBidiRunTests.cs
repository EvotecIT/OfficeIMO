using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTrackedBidiRunTests {
    private const string Text = "A \u05D0\u05D1\u05D2";

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ParagraphRetainsLogicalTextRecall(bool richRuns, bool tracking) {
        byte[] pdf = CreatePdf(richRuns, tracking);
        Assert.Equal(Text, PdfReadDocument.Open(pdf).ExtractText().Trim());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void PureRtlWordKeepsASeparateShowFromLatinPrefix(bool richRuns, bool tracking) {
        byte[] pdf = CreatePdf(richRuns, tracking);
        string raw = Encoding.ASCII.GetString(pdf);
        // This fixture has one text object. Inspect its uncompressed content,
        // excluding font streams and ToUnicode CMap strings.
        Match textObject = Regex.Match(raw, @"(?m)^BT\r?\n(?<content>[\s\S]*?)^ET\r?$", RegexOptions.CultureInvariant);
        Assert.True(textObject.Success, "The generated paragraph must contain a text object.");
        string content = textObject.Groups["content"].Value;
        int shows = Regex.Matches(content, @"\bT[jJ](?=\s|$)", RegexOptions.CultureInvariant).Count;
        Assert.True(shows >= 2,
            $"Latin prefix and visual RTL word require separate shows; richRuns={richRuns}, tracking={tracking}, shows={shows}.\n{content}");
    }

    private static byte[] CreatePdf(bool richRuns, bool tracking) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
            Text.Select(character => (int)character).Distinct().ToArray());
        if (tracking) {
            // Use the same valid tracking-table producer as the current spacing tests.
            font = ManagedTextShapingTestAssets.AddTrackingTable(font,
                ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { 0D },
                    new[] { new short[] { -100, 0 } }));
        }
        var options = new PdfOptions {
            TextShapingMode = PdfTextShapingMode.UnicodeScalar,
            CompressContentStreams = false,
            DefaultFontSize = 12D
        }.UseFontFamily(new PdfEmbeddedFontFamily("RtlIsolation", font));
        PdfDocument document = PdfDocument.Create(options);
        if (richRuns) {
            document.Paragraph(paragraph => paragraph.Runs(new[] {
                new PdfTextRun("A "), new PdfTextRun("\u05D0\u05D1\u05D2")
            }));
        } else {
            document.Paragraph(paragraph => paragraph.Text(Text));
        }
        return document.ToBytes();
    }
}

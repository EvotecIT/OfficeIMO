using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Letter = UglyToad.PdfPig.Content.Letter;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData("Footnote")]
    [InlineData("BodyPages")]
    [InlineData("BodyColumns")]
    public void JustifiedParagraphFragmentsMatchWholeParagraphGlyphsBackgroundsAndUnderlines(string owner) =>
        AssertJustificationFragmentPaint(owner, distributed: false);

    [Theory]
    [InlineData("Footnote")]
    [InlineData("BodyPages")]
    [InlineData("BodyColumns")]
    public void DistributedParagraphFragmentsMatchWholeParagraphGlyphsBackgroundsAndUnderlines(string owner) =>
        AssertJustificationFragmentPaint(owner, distributed: true);

    private static void AssertJustificationFragmentPaint(string owner, bool distributed) {
        Action<PdfContentBuilder> addParagraph = content => AddFragmentJustificationParagraph(content, distributed);
        var control = PdfDocument.Create(FragmentJustificationOptions(1000));
        addParagraph(control.Content);
        byte[] controlBytes = control.ToBytes();
        using var controlPdf = PdfPigDocument.Open(controlBytes);
        Assert.Equal(1, controlPdf.NumberOfPages);
        List<List<Letter>> expected = FragmentJustificationLines(controlPdf, columns: false);
        Assert.True(expected.Count > 10, "The paragraph must exceed one physical page or column frame.");

        PdfCreateOptions options = FragmentJustificationOptions(140);
        if (owner == "BodyColumns") options.PageWidth = 500;
        var document = PdfDocument.Create(options);
        if (owner == "Footnote") {
            PdfFootnoteDefinition note = document.Content.DefineFootnote(addParagraph);
            document.Content.Paragraph(paragraph => paragraph.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }),
                style: FragmentJustificationStyle());
        } else if (owner == "BodyColumns") {
            document.Content.Columns(addParagraph,
                new PdfMultiColumnOptions { ColumnCount = 2, Gap = 20, BalanceLastPage = false });
        } else {
            addParagraph(document.Content);
        }

        byte[] bytes = document.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        if (owner == "BodyColumns") {
            Assert.Contains(pdf.GetPage(1).Letters, letter => letter.StartBaseLine.X >= 260);
        } else {
            Assert.True(pdf.NumberOfPages >= 2, "The paragraph must continue across physical pages.");
        }
        if (owner == "Footnote") {
            Assert.Contains("Reference", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("Short", pdf.GetPage(1).Text, StringComparison.Ordinal);
        }
        List<List<Letter>> actual = FragmentJustificationLines(pdf, owner == "BodyColumns");
        Assert.Equal(expected.Count, actual.Count);
        for (int line = 0; line < expected.Count; line++) {
            Assert.Equal(string.Concat(expected[line].Select(letter => letter.Value)), string.Concat(actual[line].Select(letter => letter.Value)));
            Assert.Equal(expected[line].Count, actual[line].Count);
            for (int glyph = 0; glyph < expected[line].Count; glyph++) {
                double expectedX = expected[line][glyph].StartBaseLine.X - expected[line][0].StartBaseLine.X;
                double actualX = actual[line][glyph].StartBaseLine.X - actual[line][0].StartBaseLine.X;
                Assert.InRange(Math.Abs(expectedX - actualX), 0D, .02D);
            }
        }

        string controlContent = PdfTestContentText.ReadDecoded(controlBytes);
        string content = PdfTestContentText.ReadDecoded(bytes);
        var expectedBackgrounds = ExtractPaintedRectangles(controlContent, "1 1 0 rg", "f");
        var backgrounds = ExtractPaintedRectangles(content, "1 1 0 rg", "f");
        Assert.Equal(expected.Count, expectedBackgrounds.Count);
        Assert.Equal(expectedBackgrounds.Count, backgrounds.Count);
        for (int index = 0; index < backgrounds.Count; index++) {
            Assert.InRange(Math.Abs(expectedBackgrounds[index].W - backgrounds[index].W), 0D, .02D);
            Assert.InRange(Math.Abs(expectedBackgrounds[index].H - backgrounds[index].H), 0D, .02D);
        }
        var expectedUnderlines = ExtractStrokedLineSegments(controlContent, "0 0 0 RG");
        var underlines = ExtractStrokedLineSegments(content, "0 0 0 RG");
        Assert.NotEmpty(expectedUnderlines);
        Assert.Equal(expectedUnderlines.Count, underlines.Count);
        for (int index = 0; index < underlines.Count; index++)
            Assert.InRange(Math.Abs((expectedUnderlines[index].X2 - expectedUnderlines[index].X1) -
                (underlines[index].X2 - underlines[index].X1)), 0D, .02D);

        var expectedLinks = controlPdf.GetPages().SelectMany(page => page.GetHyperlinks()).ToArray();
        var links = pdf.GetPages().SelectMany(page => page.GetHyperlinks()).ToArray();
        Assert.NotEmpty(expectedLinks);
        Assert.Equal(expectedLinks.Length, links.Length);
        for (int index = 0; index < links.Length; index++) {
            Assert.InRange(Math.Abs(expectedLinks[index].Bounds.Width - links[index].Bounds.Width), 0D, .02D);
            Assert.InRange(Math.Abs(expectedLinks[index].Bounds.Height - links[index].Bounds.Height), 0D, .02D);
        }
    }

    private static PdfCreateOptions FragmentJustificationOptions(double height) => new() {
        PageWidth = 260, PageHeight = height, MarginLeft = 20, MarginRight = 20,
        MarginTop = 20, MarginBottom = 20, DefaultFontSize = 8, DefaultFont = PdfStandardFont.Helvetica,
        CompressContentStreams = false
    };

    private static PdfParagraphStyle FragmentJustificationStyle(bool distributed = false) => new() {
        FontSize = 8, LineSpacing = PdfLineSpacing.Exactly(10), SpacingBefore = 0, SpacingAfter = 0,
        JustificationMode = distributed ? PdfJustificationMode.Characters : PdfJustificationMode.Words,
        JustifyExplicitLineBreaks = distributed, JustifyLastLine = distributed
    };

    private static void AddFragmentJustificationParagraph(PdfContentBuilder content, bool distributed) {
        string text = "Short break\n" + string.Join(" ", Enumerable.Range(0, 80).Select(index => "word" + index.ToString("D2", CultureInfo.InvariantCulture)));
        content.Paragraph(paragraph => paragraph.Runs(new[] { new PdfTextRun(text, underline: true,
                backgroundColor: new PdfColor(1, 1, 0), linkUri: "https://example.test/fragment") }),
            PdfAlign.Justify, style: FragmentJustificationStyle(distributed));
    }

    private static List<List<Letter>> FragmentJustificationLines(PdfPigDocument pdf, bool columns) {
        var result = new List<List<Letter>>();
        foreach (var page in pdf.GetPages()) {
            for (int column = 0; column < (columns ? 2 : 1); column++) {
                double left = 20 + column * 240;
                result.AddRange(page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value) &&
                        letter.StartBaseLine.X >= left - .02D && letter.StartBaseLine.X < left + 220)
                    .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1)).OrderByDescending(group => group.Key)
                    .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
                    .Where(line => string.Concat(line.Select(letter => letter.Value)) != "Reference"));
            }
        }
        return result;
    }
}

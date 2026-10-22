using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Letter = UglyToad.PdfPig.Content.Letter;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordNativeHeadingJustificationPolicyTests {
    [Theory]
    [InlineData("both-break")]
    [InlineData("distribute-final")]
    public void SaveAsPdf_HeadingMatchesBodySourceJustificationPolicy(string scenario) {
        List<List<Letter>> natural = Render(heading: false, scenario, left: true);
        List<List<Letter>> expected = Render(heading: false, scenario, left: false);
        List<List<Letter>> actual = Render(heading: true, scenario, left: false);
        Assert.Equal(expected.Count, actual.Count);
        Assert.True(Span(expected[0]) > Span(natural[0]));
        for (int line = 0; line < expected.Count; line++) {
            Assert.Equal(string.Concat(expected[line].Select(letter => letter.Value)),
                string.Concat(actual[line].Select(letter => letter.Value)));
            for (int glyph = 0; glyph < expected[line].Count; glyph++) {
                Assert.Equal(expected[line][glyph].StartBaseLine.X - expected[line][0].StartBaseLine.X,
                    actual[line][glyph].StartBaseLine.X - actual[line][0].StartBaseLine.X, 3);
                Assert.Equal(expected[line][glyph].Width, actual[line][glyph].Width, 3);
            }
        }
    }

    private static List<List<Letter>> Render(bool heading, string scenario, bool left) {
        using WordDocument document = WordDocument.Create();
        WordParagraph target = document.AddParagraph();
        if (heading) target.SetStyle(WordParagraphStyles.Heading1);
        target._paragraph.RemoveAllChildren<W.Run>();
        target._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        target._paragraph.ParagraphProperties.Justification = new W.Justification {
            Val = left ? W.JustificationValues.Left : scenario == "both-break"
                ? W.JustificationValues.Both : W.JustificationValues.Distribute
        };
        target._paragraph.ParagraphProperties.SpacingBetweenLines = new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "280", LineRule = W.LineSpacingRuleValues.Exact
        };
        target._paragraph.Append(Run("Alpha beta gamma"));
        if (scenario == "both-break") target._paragraph.Append(new W.Run(new W.Break()), Run("Short final words"));
        Assert.Empty(document.ValidateDocument());
        string authored = document.OpenXmlDocument.MainDocumentPart!.Document.OuterXml;
        byte[] bytes = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false,
            PageSize = new PageSize(300D, 300D), Margins = PageMargins.Uniform(30D), FontFamily = "Helvetica" }).ToBytes();
        Assert.Equal(authored, document.OpenXmlDocument.MainDocumentPart.Document.OuterXml);
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-compression-evidence");
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "heading-policy-" + scenario + "-" + (heading ? "heading" : "body") + "-" + (left ? "left" : "source") + ".pdf");
        File.WriteAllBytes(path, bytes);
        using var pdf = PdfPigDocument.Open(path);
        return pdf.GetPages().SelectMany(page => page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2)).OrderByDescending(line => line.Key)
            .Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToList())).ToList();
    }

    private static W.Run Run(string text) => new(new W.RunProperties(
        new W.RunFonts { Ascii = "Helvetica", HighAnsi = "Helvetica" }, new W.FontSize { Val = "24" }),
        new W.Text(text) { Space = SpaceProcessingModeValues.Preserve });
    private static double Span(IReadOnlyList<Letter> line) => line[line.Count - 1].EndBaseLine.X - line[0].StartBaseLine.X;
}

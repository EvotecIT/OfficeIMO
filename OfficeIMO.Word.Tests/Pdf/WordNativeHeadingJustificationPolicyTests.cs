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
    [InlineData(WordParagraphAlignment.HighKashida, false)]
    [InlineData(WordParagraphAlignment.MediumKashida, false)]
    [InlineData(WordParagraphAlignment.LowKashida, false)]
    [InlineData(WordParagraphAlignment.ThaiDistribute, false)]
    [InlineData(WordParagraphAlignment.HighKashida, true)]
    [InlineData(WordParagraphAlignment.MediumKashida, true)]
    [InlineData(WordParagraphAlignment.LowKashida, true)]
    [InlineData(WordParagraphAlignment.ThaiDistribute, true)]
    public void SaveAsPdf_HeadingOutsideWordSpacingProfileRetainsItsAlignmentFallback(
        WordParagraphAlignment alignment, bool inherited) {
        List<Letter> RenderFallback(bool leftControl) {
            using WordDocument document = WordDocument.Create();
            WordParagraph target = document.AddParagraph();
            target.SetStyle(inherited ? WordParagraphStyles.Heading3 : WordParagraphStyles.Heading1);
            target._paragraph.RemoveAllChildren<W.Run>();
            WordParagraphAlignment authoredAlignment = leftControl ? WordParagraphAlignment.Left : alignment;
            if (inherited) {
                W.Style headingStyle = document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                    .Elements<W.Style>().Single(style => style.StyleId?.Value == target.StyleId);
                headingStyle.StyleParagraphProperties ??= new W.StyleParagraphProperties();
                headingStyle.StyleParagraphProperties.Justification = new W.Justification { Val = authoredAlignment.ToOpenXml() };
                target.ParagraphAlignment = null;
                Assert.Null(target.ParagraphAlignment);
            } else {
                target.ParagraphAlignment = authoredAlignment;
                Assert.Equal(authoredAlignment, target.ParagraphAlignment);
            }
            target._paragraph.Append(Run("Fallback heading words"));
            Assert.Empty(document.ValidateDocument());
            string authoredDocument = document.OpenXmlDocument.MainDocumentPart!.Document.OuterXml;
            string authoredStyles = document.OpenXmlDocument.MainDocumentPart.StyleDefinitionsPart!.Styles!.OuterXml;
            byte[] bytes = document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
                IncludePageNumbers = false, PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(300D),OfficeIMO.OfficeLength.Points(300D)),
                Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30D)), FontFamily = "Helvetica"
            }).ToBytes();
            Assert.Equal(authoredDocument, document.OpenXmlDocument.MainDocumentPart.Document.OuterXml);
            Assert.Equal(authoredStyles, document.OpenXmlDocument.MainDocumentPart.StyleDefinitionsPart.Styles.OuterXml);
            RetainEvidence(bytes, "heading-fallback-" + alignment + "-" + inherited + "-" + leftControl);
            using var pdf = PdfPigDocument.Open(bytes);
            List<Letter> letters = Assert.Single(pdf.GetPages()).Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToList();
            Assert.Equal("Fallbackheadingwords", string.Concat(letters.Select(letter => letter.Value)));
            return letters;
        }

        List<Letter> expected = RenderFallback(leftControl: true);
        List<Letter> actual = RenderFallback(leftControl: false);
        Assert.Equal(expected.Count, actual.Count);
        for (int glyph = 0; glyph < expected.Count; glyph++) {
            Assert.Equal(expected[glyph].StartBaseLine.X, actual[glyph].StartBaseLine.X, 3);
            Assert.Equal(expected[glyph].StartBaseLine.Y, actual[glyph].StartBaseLine.Y, 3);
            Assert.Equal(expected[glyph].Width, actual[glyph].Width, 3);
        }
    }

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
        byte[] bytes = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false,
            PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(300D),OfficeIMO.OfficeLength.Points(300D)), Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30D)), FontFamily = "Helvetica" }).ToBytes();
        Assert.Equal(authored, document.OpenXmlDocument.MainDocumentPart.Document.OuterXml);
        RetainEvidence(bytes, "heading-policy-" + scenario + "-" + (heading ? "heading" : "body") + "-" + (left ? "left" : "source"));
        using var pdf = PdfPigDocument.Open(bytes);
        return pdf.GetPages().SelectMany(page => page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2)).OrderByDescending(line => line.Key)
            .Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToList())).ToList();
    }

    private static void RetainEvidence(byte[] bytes, string label) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_ENGINE_PROOF_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        string directory = output!;
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, label + "-" + Guid.NewGuid().ToString("N") + ".pdf"), bytes);
    }

    private static W.Run Run(string text) => new(new W.RunProperties(
        new W.RunFonts { Ascii = "Helvetica", HighAnsi = "Helvetica" }, new W.FontSize { Val = "24" }),
        new W.Text(text) { Space = SpaceProcessingModeValues.Preserve });
    private static double Span(IReadOnlyList<Letter> line) => line[line.Count - 1].EndBaseLine.X - line[0].StartBaseLine.X;
}

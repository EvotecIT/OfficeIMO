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
public sealed class WordPdfJustificationPolicyTests {
    [Theory]
    [InlineData("absent", "body", true)]
    [InlineData("false", "body", true)]
    [InlineData("true", "body", false)]
    [InlineData("implicit", "body", false)]
    [InlineData("absent", "footnote", true)]
    [InlineData("true", "footnote", false)]
    public void SaveAsPdf_JustifiedSoftBreakRespectsSourceExpansionSetting(string sourceValue, string container, bool expands) {
        using WordDocument document = WordDocument.Create();
        var main = document.OpenXmlDocument.MainDocumentPart!;
        W.Compatibility compatibility = main.DocumentSettingsPart!.Settings!.GetFirstChild<W.Compatibility>()!;
        if (sourceValue != "absent") {
            W.DoNotExpandShiftReturn suppression = new();
            if (sourceValue != "implicit")
                suppression.Val = new OnOffValue(sourceValue == "true");
            compatibility.PrependChild(suppression);
        }
        W.Paragraph paragraph = CreateJustificationPolicyParagraph(W.JustificationValues.Both);
        if (container == "footnote") {
            document.AddParagraph("Reference").AddFootNote("Temporary");
            W.Footnote note = main.FootnotesPart!.Footnotes!.Elements<W.Footnote>()
                .Single(item => item.Type == null && item.Id?.Value > 0);
            note.RemoveAllChildren<W.Paragraph>();
            note.Append(paragraph);
        } else {
            WordParagraph target = document.AddParagraph();
            target._paragraph.InsertAfterSelf(paragraph);
            target._paragraph.Remove();
        }

        using var pdf = PdfPigDocument.Open(ConvertJustificationPolicyDocument(document));
        List<Letter> firstLine = JustificationPolicyLine(pdf, "Alphabetagamma");
        double width = firstLine.Last().EndBaseLine.X - firstLine.First().StartBaseLine.X;
        if (expands)
            Assert.InRange(width, 239.95D, 240.05D);
        else
            Assert.InRange(width, 80D, 140D);
        List<Letter> lastLine = JustificationPolicyLine(pdf, "Shortfinalwords");
        Assert.InRange(lastLine.Last().EndBaseLine.X - lastLine.First().StartBaseLine.X, 70D, 140D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_DistributedParagraphExpandsWithinWordsOnExplicitAndFinalLines(bool suppressShiftReturn) {
        using WordDocument document = WordDocument.Create();
        if (suppressShiftReturn)
            document.OpenXmlDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!
                .GetFirstChild<W.Compatibility>()!.PrependChild(new W.DoNotExpandShiftReturn { Val = true });
        WordParagraph target = document.AddParagraph();
        W.Paragraph paragraph = CreateJustificationPolicyParagraph(W.JustificationValues.Distribute);
        target._paragraph.InsertAfterSelf(paragraph);
        target._paragraph.Remove();

        using var pdf = PdfPigDocument.Open(ConvertJustificationPolicyDocument(document));
        foreach (string text in new[] { "Alphabetagamma", "Shortfinalwords" }) {
            List<Letter> line = JustificationPolicyLine(pdf, text);
            Assert.InRange(line.Last().EndBaseLine.X - line.First().StartBaseLine.X, 239.95D, 240.05D);
            // The first two glyphs belong to the same word: word spacing alone cannot satisfy this.
            Assert.True(line[1].StartBaseLine.X - line[0].EndBaseLine.X > 3D);
        }
    }

    [Fact]
    public void SaveAsPdf_TableParagraphRetainsItsOwnExplicitAndFinalLinePolicies() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Dxa, Width = "4800" };
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(4800) * 635)));
        W.Paragraph original = cell.Paragraphs[0]._paragraph;
        W.Paragraph paragraph = CreateJustificationPolicyParagraph(W.JustificationValues.Both);
        original.InsertAfterSelf(paragraph);
        original.Remove();
        W.Paragraph edge = CreateJustificationPolicyParagraph(W.JustificationValues.Right);
        edge.RemoveAllChildren<W.Run>();
        edge.Append(new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Helvetica", HighAnsi = "Helvetica" },
            new W.FontSize { Val = "24" }), new W.Text("RIGHTEDGE")));
        paragraph.InsertAfterSelf(edge);

        using var pdf = PdfPigDocument.Open(ConvertJustificationPolicyDocument(document));
        List<Letter> first = JustificationPolicyLine(pdf, "Alphabetagamma");
        List<Letter> final = JustificationPolicyLine(pdf, "Shortfinalwords");
        List<Letter> right = JustificationPolicyLine(pdf, "RIGHTEDGE");
        Assert.Equal(right.Last().EndBaseLine.X, first.Last().EndBaseLine.X, 2);
        Assert.InRange(final.Last().EndBaseLine.X - final.First().StartBaseLine.X, 70D, 140D);
        Assert.True(right.Last().EndBaseLine.X - final.Last().EndBaseLine.X > 50D,
            $"Expanded edge {right.Last().EndBaseLine.X}, source-final edge {final.Last().EndBaseLine.X}.");
    }

    private static W.Paragraph CreateJustificationPolicyParagraph(W.JustificationValues alignment) => new(
        new W.ParagraphProperties(
            new W.SpacingBetweenLines { Before = "0", After = "0", Line = "280", LineRule = W.LineSpacingRuleValues.Exact },
            new W.Justification { Val = alignment }),
        new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Helvetica", HighAnsi = "Helvetica" }, new W.FontSize { Val = "24" }),
            new W.Text("Alpha beta gamma"), new W.Break(), new W.Text("Short final words")));

    private static List<Letter> JustificationPolicyLine(PdfPigDocument pdf, string expectedText) {
        var lines = pdf.GetPages().SelectMany(page => page.Letters)
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2))
            .Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToList());
        return Assert.Single(lines, line => string.Concat(line.Select(letter => letter.Value)) == expectedText);
    }

    private static byte[] ConvertJustificationPolicyDocument(WordDocument document) {
        var main = document.OpenXmlDocument.MainDocumentPart!;
        Assert.Empty(document.ValidateDocument());
        string body = main.Document.OuterXml;
        string settings = main.DocumentSettingsPart!.Settings!.OuterXml;
        string? notes = main.FootnotesPart?.Footnotes?.OuterXml;
        byte[] bytes = document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(320)),
            Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30)),
            FontFamily = "Helvetica"
        }).ToBytes();
        Assert.Equal(body, main.Document.OuterXml);
        Assert.Equal(settings, main.DocumentSettingsPart.Settings.OuterXml);
        Assert.Equal(notes, main.FootnotesPart?.Footnotes?.OuterXml);
        Assert.Empty(document.ValidateDocument());
        return bytes;
    }
}

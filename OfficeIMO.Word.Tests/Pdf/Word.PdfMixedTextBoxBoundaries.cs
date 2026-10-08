using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MixedTextBoxRetainsSiblingChartAndGroup(bool chart) {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordParagraph paragraph;
        if (chart) {
            WordChart drawing = document.AddChart("MIXEDCHART", false, 240, 150);
            drawing.AddPie("FIRST", 4D); drawing.AddPie("SECOND", 2D);
            paragraph = document.Paragraphs.Last();
        } else {
            paragraph = document.AddParagraph("PREFIX");
            paragraph.AddShape(WordShapeType.Rectangle, 36D, 20D, "#CCB399");
            var rectangle = paragraph._paragraph.Descendants<DocumentFormat.OpenXml.Vml.Rectangle>().Single();
            var picture = Assert.IsType<W.Picture>(rectangle.Parent);
            rectangle.Remove();
            var group = new DocumentFormat.OpenXml.Vml.Group(rectangle) {
                Style = "position:absolute;left:120pt;top:100pt;width:36pt;height:20pt;z-index:-1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page",
                CoordinateSize = "36,20"
            };
            picture.Append(group);
        }
        paragraph.AddTextBox("BOX", WordImageTextWrapping.Square);
        paragraph.AddText("SUFFIX");
        Assert.Empty(document.ValidateDocument());
        var result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        byte[] bytes = result.Value.ToBytes();
        if (chart) {
            using var pdf = PdfDocument.Open(bytes);
            Assert.Contains("MIXEDCHART", string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value)));
        } else {
            string content = PdfOperatorSearchText.From(bytes);
            Assert.True(content.Contains("0.8 0.702 0.6 rg"),
                string.Join("; ", result.Report.Warnings.Select(warning => warning.Code + ":" + warning.Message)) +
                "; paints=" + string.Join(";", System.Text.RegularExpressions.Regex.Matches(content, @"[\d.]+ [\d.]+ [\d.]+ rg").Cast<System.Text.RegularExpressions.Match>().Select(match => match.Value)));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MixedTextBoxParagraphSpacingIsAppliedOnce(bool before) {
        double Measure(double spacing) {
            using WordDocument document = CreateJoinedParagraphDocument();
            WordParagraph paragraph = document.AddParagraph("PREFIX");
            paragraph.LineSpacingBeforePoints = before ? spacing : 0D;
            paragraph.LineSpacingAfterPoints = before ? 0D : spacing;
            paragraph.AddTextBox("BOX", WordImageTextWrapping.Square);
            paragraph.AddText("SUFFIX");
            document.AddParagraph("AFTER").LineSpacingBeforePoints = 0D;
            var pdf = OfficeIMO.Pdf.PdfReadDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
            return Assert.Single(pdf.Pages.SelectMany(page => page.GetTextSpans()),
                span => span.Text.Contains(before ? "PREFIX" : "AFTER")).Y;
        }
        Assert.Equal(24D, Math.Abs(Measure(24D) - Measure(0D)), precision: 3);
    }

    [Theory]
    [InlineData("body", false)]
    [InlineData("body", true)]
    [InlineData("header", false)]
    [InlineData("header", true)]
    public void MixedTextBoxEquationContentIsEmittedOnceInSourceOrder(string scope, bool equationOnlyPrefix) {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordParagraph paragraph = scope == "body" ? document.AddParagraph(equationOnlyPrefix ? "" : "PREFIX")
            : document.HeaderDefaultOrCreate.AddParagraph(equationOnlyPrefix ? "" : "PREFIX");
        paragraph.AddEquation("<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>MATH</m:t></m:r></m:oMath>");
        // A box added to the original text run precedes the following math
        // element. Use the new run after the equation to author this order.
        paragraph.AddText(string.Empty).AddTextBox("BOX", WordImageTextWrapping.Square);
        paragraph.AddText("SUFFIX");
        if (scope != "body") {
            document.HeaderDefaultOrCreate.AddParagraph().AddField(WordFieldType.SectionPages);
            document.AddParagraph("BODY");
        }
        Assert.Empty(document.ValidateDocument());
        using var pdf = PdfDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        foreach (string token in equationOnlyPrefix ? new[] { "MATH", "BOX", "SUFFIX" } : new[] { "PREFIX", "MATH", "BOX", "SUFFIX" })
            Assert.Equal(1, visible.Split(new[] { token }, StringSplitOptions.None).Length - 1);
        Assert.True(visible.IndexOf("MATH", StringComparison.Ordinal) < visible.IndexOf("BOX", StringComparison.Ordinal), visible);
        Assert.True(visible.IndexOf("BOX", StringComparison.Ordinal) < visible.IndexOf("SUFFIX", StringComparison.Ordinal), visible);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void MixedTextBoxNoteReferencesAreOwnedOnce(bool endnote, bool sharedRun) {
        using WordDocument document = CreateJoinedParagraphDocument();
        document.Sections[0].AddEndnoteProperties(WordNumberFormat.Decimal);
        WordParagraph paragraph = document.AddParagraph("PREFIX");
        paragraph.AddTextBox("BOX", WordImageTextWrapping.Square);
        paragraph.AddText("SUFFIX");
        if (endnote) paragraph.AddEndNote("NOTEBODY"); else paragraph.AddFootNote("NOTEBODY");
        if (sharedRun) {
            var runs = paragraph._paragraph.Elements<W.Run>().ToArray();
            foreach (W.Run run in runs.Skip(1)) {
                foreach (var child in run.ChildElements.Where(child => child is not W.RunProperties).ToArray()) {
                    child.Remove(); runs[0].Append(child);
                }
                run.Remove();
            }
        } else {
            // A reference-only tail is valid imported Word content; it must
            // survive even when it has no adjacent text to trigger a flush.
            DocumentFormat.OpenXml.OpenXmlElement reference = endnote
                ? paragraph._paragraph.Descendants<W.EndnoteReference>().Single()
                : paragraph._paragraph.Descendants<W.FootnoteReference>().Single();
            reference.Remove();
            paragraph._paragraph.Append(new W.Run(reference));
        }
        string before = paragraph._paragraph.OuterXml;
        Assert.Empty(document.ValidateDocument());
        using var pdf = PdfDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Equal(2, visible.Count(character => character == '1')); // Anchor and note-body label.
        Assert.Contains("NOTEBODY", visible);
        Assert.Equal(before, paragraph._paragraph.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MixedTextBoxRetainsSiblingShapePaint(bool drawingMl) {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordParagraph paragraph = document.AddParagraph("PREFIX");
        paragraph.AddTextBox("BOX", WordImageTextWrapping.Square);
        WordShape shape = drawingMl ? paragraph.AddShapeDrawing(WordShapeType.Rectangle, 36D, 20D)
            : paragraph.AddShape(WordShapeType.Rectangle, 36D, 20D, "#CCB399");
        shape.FillColorHex = "#CCB399";
        paragraph.AddText("SUFFIX");
        Assert.Empty(document.ValidateDocument());
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false });
        Assert.Contains("0.8 0.702 0.6 rg", PdfOperatorSearchText.From(bytes));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(0)]
    public void MixedTextBoxPicturesRespectTheOriginalParagraphImageLimit(int limit) {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordParagraph paragraph = document.AddParagraph("PREFIX");
        paragraph.AddTextBox("BOX", WordImageTextWrapping.Square);
        for (int index = 0; index < 2; index++) {
            using var image = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1));
            paragraph.AddImage(image, "picture" + index + ".png", 16D, 32D);
        }
        var options = new WordToPdfOptions { IncludePageNumbers = false, MaxImagesPerParagraph = limit };
        if (limit <= 0) Assert.Throws<ArgumentOutOfRangeException>(() => document.ToPdfBytes(options));
        else Assert.Throws<InvalidDataException>(() => document.ToPdfBytes(options));
    }
}

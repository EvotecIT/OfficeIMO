using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body", false)]
    [InlineData("body", true)]
    [InlineData("header", false)]
    [InlineData("header", true)]
    [InlineData("footer", false)]
    [InlineData("footer", true)]
    public void NativeFlowPreservesTextAroundPeerAndNestedBoxes(string scope, bool nested) {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordHeaderFooter? story = scope == "body" ? null : scope == "header"
            ? document.HeaderDefaultOrCreate : document.FooterDefaultOrCreate;
        WordParagraph paragraph = story == null ? document.AddParagraph("PREFIX") : story.AddParagraph("PREFIX");
        WordParagraph first = paragraph.AddTextBox("FIRSTBOX", WordImageTextWrapping.Square).Paragraphs.Single();
        if (nested) {
            first.AddTextBox("INNERBOX", WordImageTextWrapping.Square);
            first.AddText("AFTERINNER");
        }
        WordParagraph suffix = paragraph.AddText("SUFFIX");
        suffix.AddTextBox("SECONDBOX", WordImageTextWrapping.Square);
        paragraph.AddText("ENDTEXT");
        if (story != null) {
            // SECTIONPAGES deliberately selects full running-story flow even
            // when the mixed story cannot use plain numbered-story admission.
            story.AddParagraph()._paragraph.Append(new W.SimpleField(new W.Run(new W.Text("888"))) { Instruction = " SECTIONPAGES " });
            paragraph._paragraph.Append(new W.SimpleField(new W.Run(new W.Text("999"))) { Instruction = " PAGE " });
            document.AddParagraph("BODY");
        }
        string before = paragraph._paragraph.OuterXml;
        Assert.Empty(document.ValidateDocument());
        var result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeMixedTextBoxLayoutApproximated");
        using var pdf = PdfDocument.Open(result.Value.ToBytes());
        string visible = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));
        string[] tokens = nested
            ? new[] { "PREFIX", "FIRSTBOX", "INNERBOX", "AFTERINNER", "SUFFIX", "SECONDBOX", "ENDTEXT" }
            : new[] { "PREFIX", "FIRSTBOX", "SUFFIX", "SECONDBOX", "ENDTEXT" };
        int previous = -1;
        foreach (string token in tokens) {
            int position = visible.IndexOf(token, StringComparison.Ordinal);
            Assert.True(position > previous, $"Missing or reordered {token}: {visible}");
            Assert.Equal(1, visible.Split(new[] { token }, StringSplitOptions.None).Length - 1);
            previous = position;
        }
        if (story != null) {
            Assert.DoesNotContain("888", visible); Assert.DoesNotContain("999", visible);
            Assert.Equal(2, visible.Count(character => character == '1'));
        }
        Assert.Equal(before, paragraph._paragraph.OuterXml);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    public void NativeMixedTextBoxFlowRetainsSiblingInlinePictures(string scope) {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordHeaderFooter? story = scope == "body" ? null : scope == "header"
            ? document.HeaderDefaultOrCreate : document.FooterDefaultOrCreate;
        WordParagraph paragraph = story == null ? document.AddParagraph("PREFIX") : story.AddParagraph("PREFIX");
        paragraph.AddTextBox("BOX", WordImageTextWrapping.Square);
        using var image = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1));
        paragraph.AddImage(image, "inline.png", 16D, 32D);
        paragraph.AddText("SUFFIX");
        if (story != null) {
            story.AddParagraph().AddField(WordFieldType.SectionPages);
            document.AddParagraph("BODY");
        }
        Assert.Empty(document.ValidateDocument());
        using var pdf = PdfDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Single(pdf.GetPages().SelectMany(page => page.GetImages()));
        Assert.Contains("PREFIX", pdf.GetPage(1).Text); Assert.Contains("SUFFIX", pdf.GetPage(1).Text);
    }
}

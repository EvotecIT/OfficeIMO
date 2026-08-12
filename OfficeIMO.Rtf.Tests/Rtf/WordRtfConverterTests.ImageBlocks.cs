using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public void Word_Rtf_Bridge_Copies_Every_Image_Across_Pure_Image_Runs() {
        byte[] png = CreateOnePixelPng();
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        using (var stream = new MemoryStream(png, writable: false)) {
            paragraph.AddImage(stream, "first.png", 16, 16, description: "First run image");
        }
        var secondRun = new WordParagraph(word, paragraph._paragraph, newRun: true);
        using (var stream = new MemoryStream(png, writable: false)) {
            secondRun.AddImage(stream, "second.png", 16, 16, description: "Second run image");
        }

        RtfConversionResult<RtfDocument> conversion = word.ToRtfDocumentResult();

        Assert.Collection(
            conversion.Value.Blocks,
            block => Assert.Equal("First run image", Assert.IsType<RtfImage>(block).Description),
            block => Assert.Equal("Second run image", Assert.IsType<RtfImage>(block).Description));
        Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "WordRtfImagesOmitted");
        Assert.Same(conversion.Value, conversion.RequireNoLoss());
    }

    [Fact]
    public void Word_Rtf_Bridge_Does_Not_Use_Image_Block_Shortcut_When_A_Later_Run_Has_Text() {
        byte[] png = CreateOnePixelPng();
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        using (var stream = new MemoryStream(png, writable: false)) {
            paragraph.AddImage(stream, "inline.png", 16, 16, description: "Inline image");
        }
        new WordParagraph(word, paragraph._paragraph, newRun: true).AddText("After image");

        RtfConversionResult<RtfDocument> conversion = word.ToRtfDocumentResult();

        RtfParagraph converted = Assert.Single(conversion.Value.Paragraphs);
        Assert.Collection(
            converted.Inlines,
            inline => Assert.Equal("Inline image", Assert.IsType<RtfImage>(inline).Description),
            inline => Assert.Equal("After image", Assert.IsType<RtfRun>(inline).Text));
        Assert.Same(conversion.Value, conversion.RequireNoLoss());
    }
}

using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void TransformTextCasePreservesRunFormatting() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "TextCase.docx"));
        WordParagraph run = document.AddParagraph("Styled");
        run.SetBold().SetItalic().SetUnderline(WordUnderlineStyle.WavyDouble)
            .SetDoubleStrike().SetSuperScript().SetColor(OfficeColor.FromRgb(51, 102, 153)).SetFontFamily("Aptos");

        run.TransformTextCase(OfficeTextCase.ToggleCase);

        WordParagraph actual = document.Paragraphs.Single();
        Assert.Equal("sTYLED", actual.Text);
        Assert.True(actual.Bold);
        Assert.True(actual.Italic);
        Assert.Equal(WordUnderlineStyle.WavyDouble, actual.Underline);
        Assert.True(actual.DoubleStrike);
        Assert.Equal(WordVerticalTextPosition.Superscript, actual.VerticalTextAlignment);
        Assert.Equal("Aptos", actual.FontFamily);
    }
}

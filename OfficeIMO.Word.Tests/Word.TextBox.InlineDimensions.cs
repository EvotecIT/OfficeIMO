using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextBoxDimensionsRemainEditableAcrossInlineConversionAndDocxRoundTrip(bool inline) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Host");
        WordTextBox box = paragraph.AddTextBox("Editable", WordImageTextWrapping.Square);
        box.WidthCentimeters = 2.5D;
        box.HeightCentimeters = 1.5D;
        WordTextBox sibling = paragraph.AddTextBox("Sibling", WordImageTextWrapping.Square);
        sibling.WidthCentimeters = 2D;
        sibling.HeightCentimeters = 1D;
        if (inline) box.WrapText = WordImageTextWrapping.InLineWithText;
        Assert.Equal(2.5D, box.WidthCentimeters, 6);
        Assert.Equal(1.5D, box.HeightCentimeters, 6);
        box.WidthCentimeters = 4.5D;
        box.HeightCentimeters = 3.5D;
        Assert.Equal(4.5D, box.WidthCentimeters, 6);
        Assert.Equal(3.5D, box.HeightCentimeters, 6);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        WordTextBox edited = reopened.TextBoxes[0];
        Assert.Equal(inline ? WordImageTextWrapping.InLineWithText : WordImageTextWrapping.Square, edited.WrapText);
        Assert.Equal(4.5D, edited.WidthCentimeters, 6);
        Assert.Equal(3.5D, edited.HeightCentimeters, 6);
        Assert.Equal(2D, reopened.TextBoxes[1].WidthCentimeters, 6);
        Assert.Equal(1D, reopened.TextBoxes[1].HeightCentimeters, 6);
        Assert.Empty(reopened.ValidateDocument());
    }
}

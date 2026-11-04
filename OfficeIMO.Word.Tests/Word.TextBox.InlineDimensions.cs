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
        box.Width = OfficeLength.Centimeters(2.5D);
        box.RelativeWidthPercentage = 0;
        box.Height = OfficeLength.Centimeters(1.5D);
        box.RelativeHeightPercentage = 0;
        WordTextBox sibling = paragraph.AddTextBox("Sibling", WordImageTextWrapping.Square);
        sibling.Width = OfficeLength.Centimeters(2D);
        sibling.RelativeWidthPercentage = 0;
        sibling.Height = OfficeLength.Centimeters(1D);
        sibling.RelativeHeightPercentage = 0;
        if (inline) box.WrapText = WordImageTextWrapping.InLineWithText;
        Assert.Equal(2.5D, box.Width.ToCentimeters(), 6);
        Assert.Equal(1.5D, box.Height.ToCentimeters(), 6);
        box.Width = OfficeLength.Centimeters(4.5D);
        box.RelativeWidthPercentage = 0;
        box.Height = OfficeLength.Centimeters(3.5D);
        box.RelativeHeightPercentage = 0;
        Assert.Equal(4.5D, box.Width.ToCentimeters(), 6);
        Assert.Equal(3.5D, box.Height.ToCentimeters(), 6);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        WordTextBox edited = reopened.TextBoxes[0];
        Assert.Equal(inline ? WordImageTextWrapping.InLineWithText : WordImageTextWrapping.Square, edited.WrapText);
        Assert.Equal(4.5D, edited.Width.ToCentimeters(), 6);
        Assert.Equal(3.5D, edited.Height.ToCentimeters(), 6);
        Assert.Equal(2D, reopened.TextBoxes[1].Width.ToCentimeters(), 6);
        Assert.Equal(1D, reopened.TextBoxes[1].Height.ToCentimeters(), 6);
        Assert.Empty(reopened.ValidateDocument());
    }
}

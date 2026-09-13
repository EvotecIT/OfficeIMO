using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void TableBorderHelpersSaveSchemaValidBorders(bool separateInside) {
            using var document = WordDocument.Create();
            var table = document.AddTable(2, 2, WordTableStyle.TableGrid);
            table.SetCellText(0, 0, "Record"); table.SetCellText(1, 0, "Example");
            if (separateInside) table.StyleDetails!.SetBordersOutsideInside(WordBorderStyle.Single, 8, OfficeColor.Blue, WordBorderStyle.Single, 4, OfficeColor.Red);
            else table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Single, 8, OfficeColor.Blue);
            using var saved = WordDocument.Load(new MemoryStream(document.ToBytes()));
            Assert.Empty(saved.ValidateDocument());
            Assert.Equal("0000FF", saved.Tables[0].StyleDetails!.GetBorderProperties(WordTableBorderSide.Top).ColorHex);
            Assert.Equal(separateInside ? "FF0000" : "0000FF", saved.Tables[0].StyleDetails!.GetBorderProperties(WordTableBorderSide.InsideHorizontal).ColorHex);
        }
    }
}

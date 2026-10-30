using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData(2, false)]
        [InlineData(257, true)]
        public void TableLayoutStyleChainResolvesOrRejectsBoundedDepth(int styleCount, bool exceedsLimit) {
            using WordDocument document = WordDocument.Create();
            Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            for (int index = 0; index < styleCount; index++) {
                var style = new Style { Type = StyleValues.Table, StyleId = "Layout" + index };
                if (index > 0) style.Append(new BasedOn { Val = "Layout" + (index - 1) });
                else style.Append(new StyleTableProperties(new TableLayout { Type = TableLayoutValues.Fixed }));
                styles.Append(style);
            }
            WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
            table._tableProperties!.TableStyle = new TableStyle { Val = "Layout" + (styleCount - 1) };

            if (exceedsLimit) Assert.Throws<InvalidDataException>(() => table.LayoutMode);
            else Assert.Equal(WordTableLayoutMode.Fixed, table.LayoutMode);
        }
    }
}

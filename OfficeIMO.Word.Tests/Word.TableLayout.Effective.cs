using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("directStyle")]
    [InlineData("basedOn")]
    [InlineData("defaultStyle")]
    public void TableLayoutModeReadsInheritedFixedWithoutChangingSource(string source) {
        using WordDocument document = WordDocument.Create();
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        var fixedStyle = new W.Style {
            Type = W.StyleValues.Table, StyleId = "FixedLayoutContract", CustomStyle = true
        };
        fixedStyle.Append(new W.StyleName { Val = "Fixed layout contract" });
        fixedStyle.Append(new W.StyleTableProperties(new W.TableLayout { Type = W.TableLayoutValues.Fixed }));
        styles.Append(fixedStyle);
        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableLayout = null;
        if (source == "basedOn") {
            var derivedStyle = new W.Style { Type = W.StyleValues.Table, StyleId = "DerivedLayoutContract", CustomStyle = true };
            derivedStyle.Append(new W.StyleName { Val = "Derived layout contract" });
            derivedStyle.Append(new W.BasedOn { Val = fixedStyle.StyleId });
            styles.Append(derivedStyle);
            table._tableProperties.TableStyle = new W.TableStyle { Val = derivedStyle.StyleId };
        } else if (source == "defaultStyle") {
            foreach (W.Style style in styles.Elements<W.Style>().Where(style => style.Type?.Value == W.StyleValues.Table))
                style.Default = false;
            fixedStyle.Default = true;
            table._tableProperties.TableStyle = null;
        } else {
            table._tableProperties.TableStyle = new W.TableStyle { Val = fixedStyle.StyleId };
        }
        string tableXml = table._table.OuterXml;
        string stylesXml = styles.OuterXml;
        Assert.Equal(WordTableLayoutMode.Fixed, table.LayoutMode);
        Assert.Equal(tableXml, table._table.OuterXml);
        Assert.Equal(stylesXml, styles.OuterXml);
        table.LayoutMode = WordTableLayoutMode.AutoFit;
        Assert.Equal(WordTableLayoutMode.AutoFit, table.LayoutMode);
    }
}

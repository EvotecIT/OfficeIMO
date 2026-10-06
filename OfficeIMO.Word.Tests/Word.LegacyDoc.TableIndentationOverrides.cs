using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(-108, false)]
    [InlineData(-108, true)]
    [InlineData(720, false)]
    [InlineData(720, true)]
    public void LegacyDoc_TableIndentation_ExplicitZeroOverridesDirectOrBaseStyle(int baseIndent, bool derivedStyle) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        var baseStyle = new Style { Type = StyleValues.Table, StyleId = "IndentedTable", CustomStyle = true };
        baseStyle.Append(new StyleName { Val = "Indented Table" });
        baseStyle.Append(new BasedOn { Val = "TableNormal" });
        baseStyle.Append(new StyleTableProperties(new TableIndentation {
            Width = baseIndent, Type = TableWidthUnitValues.Dxa
        }));
        styles.Append(baseStyle);

        if (derivedStyle) {
            var childStyle = new Style { Type = StyleValues.Table, StyleId = "ZeroTable", CustomStyle = true };
            childStyle.Append(new StyleName { Val = "Zero Table" });
            childStyle.Append(new BasedOn { Val = baseStyle.StyleId });
            childStyle.Append(new StyleTableProperties(new TableIndentation {
                Width = 0, Type = TableWidthUnitValues.Dxa
            }));
            styles.Append(childStyle);
        }

        WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
        table._tableProperties!.TableStyle = new TableStyle { Val = derivedStyle ? "ZeroTable" : "IndentedTable" };
        if (!derivedStyle) table.StyleDetails!.TableIndentationWidth = 0;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Explicit zero indentation";

        using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        WordTable reloaded = Assert.Single(reopened.Tables);
        // The native zero edge may omit an optional property, but its effective indent must remain zero.
        Assert.Equal((short)0, reloaded.StyleDetails!.TableIndentationWidth.GetValueOrDefault());
    }
}

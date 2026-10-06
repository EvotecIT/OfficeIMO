using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_IndependentTableConditionsPreserveDerivedParagraphSpacing(bool alignmentFirst) {
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(new StyleName { Val = "Base row spacing" },
            new TableStyleProperties(new StyleParagraphProperties(new SpacingBetweenLines {
                Line = "360", LineRule = LineSpacingRuleValues.Exact, After = "80"
            })) { Type = TableStyleOverrideValues.FirstRow }) {
            Type = StyleValues.Table, StyleId = "BaseRowSpacing", CustomStyle = true
        });
        var rowCondition = new TableStyleProperties(new StyleParagraphProperties(new SpacingBetweenLines {
            Line = "420", LineRule = LineSpacingRuleValues.Auto
        })) { Type = TableStyleOverrideValues.FirstRow };
        var columnCondition = new TableStyleProperties(new StyleParagraphProperties(
            new Justification { Val = JustificationValues.Center })) { Type = TableStyleOverrideValues.FirstColumn };
        var derived = new Style(new StyleName { Val = "Derived row spacing with column alignment" },
            new BasedOn { Val = "BaseRowSpacing" }) {
            Type = StyleValues.Table, StyleId = "DerivedRowSpacing", CustomStyle = true
        };
        if (alignmentFirst) derived.Append(columnCondition, rowCondition);
        else derived.Append(rowCondition, columnCondition);
        styles.Append(derived);
        WordTable table = source.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "DerivedRowSpacing" };
        table.ConditionalFormattingFirstRow = true;
        table.ConditionalFormattingFirstColumn = true;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Inherited";
        string originalStyles = styles.OuterXml;
        string originalBody = source._wordprocessingDocument.MainDocumentPart.Document.OuterXml;

        byte[] bytes = source.ToBytes(WordFileFormat.Doc);

        Assert.Equal(originalStyles, styles.OuterXml);
        Assert.Equal(originalBody, source._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        WordParagraph paragraph = reopened.Tables.Single().Rows[0].Cells[0].Paragraphs[0];
        Assert.Equal("Inherited", paragraph.Text);
        Assert.Equal(420, paragraph.LineSpacing);
        Assert.Equal(WordLineSpacingRule.Auto, paragraph.LineSpacingRule);
        Assert.Equal(4, paragraph.LineSpacingAfterPoints);
        Assert.Equal(JustificationValues.Center, paragraph._paragraph.ParagraphProperties!.Justification!.Val!.Value);
    }
}

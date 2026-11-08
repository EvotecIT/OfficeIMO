using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LegacyDoc_SavePreservesCaseDistinctParagraphAndTableStyleInheritance() {
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(
            new Style(
                new StyleName { Val = "Uppercase paragraph" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new StyleRunProperties(new Bold())) {
                Type = StyleValues.Paragraph, StyleId = "NativeDocCase", CustomStyle = true
            },
            new Style(
                new StyleName { Val = "Lowercase paragraph" },
                new BasedOn { Val = "NativeDocCase" },
                new StyleParagraphProperties(new Justification { Val = JustificationValues.Right }),
                new StyleRunProperties(new Italic())) {
                Type = StyleValues.Paragraph, StyleId = "nativeDocCase", CustomStyle = true
            },
            new Style(
                new StyleName { Val = "Uppercase table" },
                new BasedOn { Val = "TableNormal" },
                new StyleTableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Fill = "00FF00" })) {
                Type = StyleValues.Table, StyleId = "NativeDocTableCase", CustomStyle = true
            },
            new Style(
                new StyleName { Val = "Lowercase table" },
                new BasedOn { Val = "NativeDocTableCase" },
                new StyleTableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Fill = "FFFF00" })) {
                Type = StyleValues.Table, StyleId = "nativeDocTableCase", CustomStyle = true
            });
        foreach (string styleId in new[] { "NativeDocCase", "nativeDocCase" }) {
            source.AddParagraph(styleId)._paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = styleId });
        }
        foreach (string styleId in new[] { "NativeDocTableCase", "nativeDocTableCase" }) {
            WordTable table = source.AddTable(1, 1, WordTableStyle.TableNormal);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.Rows[0].Cells[0].Paragraphs[0].Text = styleId;
        }
        Assert.Empty(source.ValidateDocument());
        string authoredStyles = styles.OuterXml;

        using var native = new MemoryStream();
        source.Save(native, new WordSaveOptions { Format = WordFileFormat.Doc });
        native.Position = 0;
        using WordDocument restored = WordDocument.Load(native, options: null, out OfficeConversionReport report);

        Assert.False(report.HasLoss);
        Assert.Equal(WordFileFormat.Doc, restored.SourceFormat);
        Styles imported = restored._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        WordParagraph upper = Assert.Single(restored.Paragraphs, paragraph => paragraph.Text == "NativeDocCase");
        WordParagraph lower = Assert.Single(restored.Paragraphs, paragraph => paragraph.Text == "nativeDocCase");
        Style upperStyle = Assert.Single(imported.Elements<Style>(), style => style.StyleId == upper.StyleId);
        Style lowerStyle = Assert.Single(imported.Elements<Style>(), style => style.StyleId == lower.StyleId);
        Assert.Equal(new[] { "Uppercase paragraph", "Lowercase paragraph", "00FF00", "FFFF00" }, new[] {
            upperStyle.StyleName!.Val!.Value,
            lowerStyle.StyleName!.Val!.Value,
            restored.Tables[0].Rows[0].Cells[0].ShadingFillColor?.ToRgbHex(),
            restored.Tables[1].Rows[0].Cells[0].ShadingFillColor?.ToRgbHex()
        });
        Assert.Equal(upper.StyleId, lowerStyle.BasedOn!.Val!.Value);
        Assert.NotNull(upperStyle.StyleRunProperties!.Bold);
        Assert.Null(upperStyle.StyleRunProperties.Italic);
        Assert.Null(lowerStyle.StyleRunProperties!.Bold);
        Assert.NotNull(lowerStyle.StyleRunProperties.Italic);
        Assert.Equal(JustificationValues.Center, upperStyle.StyleParagraphProperties!.Justification!.Val!.Value);
        Assert.Equal(JustificationValues.Right, lowerStyle.StyleParagraphProperties!.Justification!.Val!.Value);
        Assert.Equal(authoredStyles, styles.OuterXml);
        Assert.Empty(restored.ValidateDocument());
    }
}

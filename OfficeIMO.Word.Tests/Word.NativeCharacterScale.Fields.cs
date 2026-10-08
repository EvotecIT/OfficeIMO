using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, 100)]
    [InlineData(true, 100)]
    [InlineData(false, 150)]
    [InlineData(true, 150)]
    public void NativeDocCharacterScaleComparesFieldRunsUsingTheParagraphRootDefault(bool complex, int explicitWidth) {
        for (int story = 0; story < 4; story++) {
            using WordDocument source = WordDocument.Create();
            Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.DocDefaults ??= new DocDefaults();
            styles.DocDefaults.RunPropertiesDefault ??= new RunPropertiesDefault(new RunPropertiesBaseStyle());
            styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle!.AddChild(new CharacterScale { Val = 150L }, true);
            Style normal = styles.Elements<Style>().Single(style => style.StyleId == "Normal");
            normal.StyleRunProperties ??= new StyleRunProperties();
            normal.StyleRunProperties.AddChild(new CharacterScale { Val = 100L }, true);
            const string rootId = "FieldWidthRoot";
            styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = rootId, CustomStyle = true,
                StyleName = new StyleName { Val = rootId } });
            Paragraph[] paragraphs = CreateFieldFormattingStories(source);
            foreach (Paragraph paragraph in new[] { paragraphs[0], paragraphs[story] }.Distinct()) {
                paragraph.ParagraphProperties ??= new ParagraphProperties();
                paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = rootId };
            }
            AppendSplitFormattingField(paragraphs[story], complex, true, "FIRST", "SECOND");
            Run second = paragraphs[story].Descendants<Run>().Single(run => run.InnerText == "SECOND");
            second.RunProperties!.AddChild(new CharacterScale { Val = explicitWidth }, true);
            Assert.Empty(source.ValidateDocument());

            if (explicitWidth == 100) {
                NotSupportedException error = Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
                Assert.Contains("display runs use one formatting set", error.Message);
            } else {
                using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
                MainDocumentPart main = loaded._wordprocessingDocument.MainDocumentPart!;
                SimpleField field = Assert.Single(EnumerateFieldFormattingRoots(main).SelectMany(root => root.Descendants<SimpleField>()));
                Assert.Equal("FIRSTSECOND", field.InnerText);
                Assert.Empty(loaded.ValidateDocument());
            }
        }
    }

    [Theory]
    [InlineData(false, 100)]
    [InlineData(true, 100)]
    [InlineData(false, 150)]
    [InlineData(true, 150)]
    public void NativeDocCharacterScaleAcceptsEquivalentFieldDefaultsAcrossStories(bool complex, int width) {
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        if (width != 100) {
            styles.DocDefaults ??= new DocDefaults();
            styles.DocDefaults.RunPropertiesDefault ??= new RunPropertiesDefault(new RunPropertiesBaseStyle());
            styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle!.AddChild(new CharacterScale { Val = width }, true);
        }
        Paragraph[] paragraphs = CreateFieldFormattingStories(source);
        for (int index = 0; index < paragraphs.Length; index++) {
            AppendSplitFormattingField(paragraphs[index], complex, true, "FIRST" + index, "SECOND" + index);
            Run second = paragraphs[index].Descendants<Run>().Single(run => run.InnerText == "SECOND" + index);
            second.RunProperties!.AddChild(new CharacterScale { Val = width }, true);
        }
        string originalStyles = styles.OuterXml;
        Assert.Empty(source.ValidateDocument());
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        SimpleField[] fields = EnumerateFieldFormattingRoots(loaded._wordprocessingDocument.MainDocumentPart!)
            .SelectMany(root => root.Descendants<SimpleField>()).ToArray();
        for (int index = 0; index < paragraphs.Length; index++) {
            Assert.Contains(fields, field => field.InnerText == "FIRST" + index + "SECOND" + index);
        }
        Assert.Equal(originalStyles, styles.OuterXml);
        Assert.Empty(loaded.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeDocCharacterScaleFieldComparisonKeepsTableStylePrecedence(bool complex) {
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults ??= new DocDefaults();
        styles.DocDefaults.RunPropertiesDefault ??= new RunPropertiesDefault(new RunPropertiesBaseStyle());
        styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle!.AddChild(new CharacterScale { Val = 150L }, true);
        const string tableStyleId = "FieldWidthTable";
        styles.Append(new Style {
            StyleId = tableStyleId, Type = StyleValues.Table, CustomStyle = true,
            StyleName = new StyleName { Val = tableStyleId },
            StyleRunProperties = new StyleRunProperties(new CharacterScale { Val = 125L })
        });
        WordTable table = source.AddTable(1, 1);
        table._table.GetFirstChild<TableProperties>()!.TableStyle = new TableStyle { Val = tableStyleId };
        Paragraph paragraph = table.Rows[0].Cells[0].Paragraphs[0]._paragraph;
        AppendSplitFormattingField(paragraph, complex, true, "FIRST", "SECOND");
        paragraph.Descendants<Run>().Single(run => run.InnerText == "SECOND").RunProperties!
            .AddChild(new CharacterScale { Val = 125L }, true);
        Assert.Empty(source.ValidateDocument());
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        SimpleField field = Assert.Single(loaded._wordprocessingDocument.MainDocumentPart!.Document.Descendants<SimpleField>());
        Assert.Equal("FIRSTSECOND", field.InnerText);
        Assert.All(field.Descendants<Run>(), run => Assert.Equal(125L, run.RunProperties!.GetFirstChild<CharacterScale>()!.Val!.Value));
        Assert.Empty(loaded.ValidateDocument());
    }
}

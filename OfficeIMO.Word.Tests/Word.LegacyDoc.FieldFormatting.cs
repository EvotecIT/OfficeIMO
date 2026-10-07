using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LegacyDoc_FieldRunsWithEquivalentFormattingRemainSupported(bool complex, bool automaticColor) {
        using WordDocument source = WordDocument.Create();
        W.Paragraph[] paragraphs = CreateFieldFormattingStories(source);
        for (int index = 0; index < paragraphs.Length; index++) {
            AppendSplitFormattingField(paragraphs[index], complex, automaticColor, "FIRST" + index, "SECOND" + index);
        }
        Assert.Empty(source.ValidateDocument());

        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        MainDocumentPart main = reopened._wordprocessingDocument.MainDocumentPart!;
        W.SimpleField[] fields = EnumerateFieldFormattingRoots(main).SelectMany(root => root.Descendants<W.SimpleField>()).ToArray();
        for (int index = 0; index < paragraphs.Length; index++) {
            Assert.Contains(fields, field => field.InnerText == "FIRST" + index + "SECOND" + index);
        }
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_FieldRunsWithDifferentInheritedVisibilityFailBeforeWriting(bool complex) {
        for (int story = 0; story < 6; story++) {
            using WordDocument source = WordDocument.Create();
            W.Style normal = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
            normal.StyleRunProperties ??= new W.StyleRunProperties();
            normal.StyleRunProperties.Vanish = new W.Vanish();
            AppendSplitFormattingField(CreateFieldFormattingStories(source)[story], complex, false, "HIDDEN", "VISIBLE");
            Assert.Empty(source.ValidateDocument());
            NotSupportedException error = Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
            Assert.Contains("display runs use one formatting set", error.Message);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_FieldRunsRespectInheritedTableVisibility(bool complex) {
        using WordDocument source = WordDocument.Create();
        const string styleId = "HiddenFieldTable";
        source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new W.Style {
            StyleId = styleId, Type = W.StyleValues.Table, CustomStyle = true,
            StyleName = new W.StyleName { Val = styleId },
            StyleRunProperties = new W.StyleRunProperties(new W.Vanish())
        });
        WordTable table = source.AddTable(1, 1);
        table._table.GetFirstChild<W.TableProperties>()!.TableStyle = new W.TableStyle { Val = styleId };
        AppendSplitFormattingField(table.Rows[0].Cells[0].Paragraphs[0]._paragraph, complex, false, "HIDDEN", "VISIBLE");
        Assert.Empty(source.ValidateDocument());
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
        Assert.Contains("display runs use one formatting set", error.Message);
    }

    private static W.Paragraph[] CreateFieldFormattingStories(WordDocument source) {
        source.AddHeadersAndFooters();
        WordParagraph notes = source.AddParagraph("References");
        W.Paragraph footnote = notes.AddFootNote("Placeholder").FootNote!.Paragraphs![1]._paragraph;
        W.Paragraph endnote = notes.AddEndNote("Placeholder").EndNote!.Paragraphs![1]._paragraph;
        return new[] {
            source.AddParagraph()._paragraph, source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0]._paragraph,
            source.Sections[0].Header.Default!.AddParagraph()._paragraph,
            source.Sections[0].Footer.Default!.AddParagraph()._paragraph, footnote, endnote
        };
    }

    private static IEnumerable<OpenXmlPartRootElement> EnumerateFieldFormattingRoots(MainDocumentPart main) {
        yield return main.Document!;
        foreach (HeaderPart header in main.HeaderParts) yield return header.Header!;
        foreach (FooterPart footer in main.FooterParts) yield return footer.Footer!;
        yield return main.FootnotesPart!.Footnotes!;
        yield return main.EndnotesPart!.Endnotes!;
    }

    private static void AppendSplitFormattingField(W.Paragraph paragraph, bool complex, bool automaticColor, string first, string second) {
        paragraph.RemoveAllChildren<W.Run>();
        var secondProperties = automaticColor
            ? new W.RunProperties(new W.Color { Val = "auto" })
            : new W.RunProperties(new W.Vanish { Val = false });
        var firstRun = new W.Run(new W.Text(first));
        var secondRun = new W.Run(secondProperties, new W.Text(second));
        if (complex) {
            paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
                new W.Run(new W.FieldCode(" DOCPROPERTY Title ")),
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
                firstRun, secondRun, new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        } else {
            paragraph.Append(new W.SimpleField(firstRun, secondRun) { Instruction = " DOCPROPERTY Title " });
        }
    }
}

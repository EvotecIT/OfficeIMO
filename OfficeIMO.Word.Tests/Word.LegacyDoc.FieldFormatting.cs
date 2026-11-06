using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using OfficeIMO.Core.Internal;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void LegacyDoc_FieldCaches_PreserveFormattingAcrossStories(int story) {
        using var input = new MemoryStream(LegacyDocTestBuilder.CreateFieldCacheCarrier(story));
        using WordDocument imported = WordDocument.Load(input);
        using var output = new MemoryStream();
        imported.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        MainDocumentPart main = reopened._wordprocessingDocument.MainDocumentPart!;
        W.SimpleField field = Assert.Single(EnumerateFieldFormattingRoots(main).SelectMany(root => root.Descendants<W.SimpleField>()));
        W.Run first = Assert.Single(field.Descendants<W.Run>(), run => run.InnerText == "FIRST");
        W.Run second = Assert.Single(field.Descendants<W.Run>(), run => run.InnerText == "SECOND");
        Assert.Null(first.RunProperties?.Bold);
        Assert.NotNull(second.RunProperties?.Bold);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void LegacyDoc_FieldCaches_PreserveOverlappingRangePriority() {
        using var input = new MemoryStream(LegacyDocTestBuilder.CreateFieldCacheCarrier(0, overlapping: true));
        using WordDocument imported = WordDocument.Load(input);
        using var output = new MemoryStream();
        imported.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        W.SimpleField field = Assert.Single(reopened._wordprocessingDocument.MainDocumentPart!.Document.Descendants<W.SimpleField>());
        Assert.Equal("FIRSTSECOND", field.InnerText);
        Assert.NotEmpty(field.Descendants<W.Italic>());
        Assert.Empty(field.Descendants<W.Bold>());
    }

    [Theory]
    [InlineData(" FILENAME ", "old.doc")]
    [InlineData(" STYLEREF \"Heading 1\" ", "Heading")]
    public void LegacyDoc_FieldTextReferences_SurviveNativeSaveAndEditableMigration(string instruction, string cachedText) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Heading").Style = WordParagraphStyles.Heading1;
        source.AddParagraph()._paragraph.Append(new W.SimpleField(new W.Run(new W.Text(cachedText))) { Instruction = instruction });
        source.PageSetup.Header.AddParagraph()._paragraph.Append(new W.SimpleField(new W.Run(new W.Text(cachedText))) { Instruction = instruction });
        using var native = new MemoryStream(source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }));
        using WordDocument imported = WordDocument.Load(native);
        using var output = new MemoryStream();
        imported.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        MainDocumentPart main = reopened._wordprocessingDocument.MainDocumentPart!;
        W.SimpleField body = Assert.Single(main.Document.Descendants<W.SimpleField>());
        W.SimpleField header = Assert.Single(main.HeaderParts.SelectMany(part => part.Header.Descendants<W.SimpleField>()));
        Assert.Equal(instruction, body.Instruction?.Value);
        Assert.Equal(cachedText, body.InnerText);
        Assert.Equal(instruction, header.Instruction?.Value);
        Assert.Equal(cachedText, header.InnerText);
        Assert.NotEmpty(imported.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }));
    }

    [Fact]
    public void LegacyDoc_FieldComparisonRefreshesStylesBetweenSaves() {
        using WordDocument source = WordDocument.Create();
        AppendSplitFormattingField(source.AddParagraph()._paragraph, complex: false, automaticColor: false, "FIRST", "SECOND");
        source.ToBytes(WordFileFormat.Doc);

        W.Style normal = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleRunProperties ??= new W.StyleRunProperties();
        normal.StyleRunProperties.Vanish = new W.Vanish();

        NotSupportedException error = Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
        Assert.Contains("display runs use one formatting set", error.Message);
    }

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

        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc })));
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
            NotSupportedException error = Assert.Throws<NotSupportedException>(() => source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }));
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
        WordTable table = source.AddTable(1, 1, WordTableStyle.TableGrid);
        table._table.GetFirstChild<W.TableProperties>()!.TableStyle = new W.TableStyle { Val = styleId };
        AppendSplitFormattingField(table.Rows[0].Cells[0].Paragraphs[0]._paragraph, complex, false, "HIDDEN", "VISIBLE");
        Assert.Empty(source.ValidateDocument());
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }));
        Assert.Contains("display runs use one formatting set", error.Message);
    }

    private static W.Paragraph[] CreateFieldFormattingStories(WordDocument source) {
        _ = source.PageSetup.Header;
        WordParagraph notes = source.AddParagraph("References");
        W.Paragraph footnote = notes.AddFootNote("Placeholder").Paragraph.FootNote!.Paragraphs![1]._paragraph;
        W.Paragraph endnote = notes.AddEndNote("Placeholder").Paragraph.EndNote!.Paragraphs![1]._paragraph;
        return new[] {
            source.AddParagraph()._paragraph, source.AddTable(1, 1, WordTableStyle.TableGrid).Rows[0].Cells[0].Paragraphs[0]._paragraph,
            source.Sections[0].PageSetup.Header.AddParagraph()._paragraph,
            source.Sections[0].PageSetup.Footer.AddParagraph()._paragraph, footnote, endnote
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

    private static partial class LegacyDocTestBuilder {
        internal static byte[] CreateFieldCacheCarrier(int story, bool overlapping = false) {
            const string field = "\u0013 FILENAME \u0014FIRSTSECOND\u0015";
            byte[] input = story switch {
                0 => CreateSimpleDoc(field),
                1 => CreateSimpleDocWithHeaderFooterStories("Body", field, string.Empty),
                2 => CreateSimpleDocWithFootnoteStory("Body", field),
                _ => CreateSimpleDocWithEndnoteStory("Body", field)
            };
            OfficeCompoundFileReader.TryRead(input, out OfficeCompoundFile? compound, out _);
            byte[] main = compound!.Streams["WordDocument"];
            byte[] table = compound.Streams["1Table"];
            int textStart = (int)((uint)BitConverter.ToInt32(table, 15) & 0x3FFFFFFF) / 2;
            int textEnd = textStart + BitConverter.ToInt32(table, 9);
            int cache = System.Text.Encoding.ASCII.GetString(main).IndexOf("FIRSTSECOND", StringComparison.Ordinal);
            int page = (main.Length + 511) & ~511;
            Array.Resize(ref main, page + (overlapping ? 1024 : 512));
            int bin = table.Length;
            Array.Resize(ref table, bin + (overlapping ? 20 : 12));
            WriteInt32(table, bin, textStart);
            WriteInt32(table, bin + 4, overlapping ? cache + 5 : textEnd);
            if (overlapping) {
                WriteInt32(table, bin + 8, textEnd);
                WriteInt32(table, bin + 12, page / 512);
                WriteInt32(table, bin + 16, page / 512 + 1);
                WriteChpxFkp(main, page, new[] { cache, cache + 11 }, -1, 0);
                WriteChpxFkp(main, page + 512, new[] { cache + 5, cache + 11 }, 0, -1);
            } else {
                WriteInt32(table, bin + 8, page / 512);
                WriteChpxFkp(main, page, new[] { textStart, cache, cache + 5, cache + 11, textEnd }, 2, -1);
            }
            WriteInt32(main, 0xFA, bin);
            WriteInt32(main, 0xFE, overlapping ? 20 : 12);
            return OfficeCompoundFileWriter.Write(new[] {
                new OfficeCompoundStream("WordDocument", main), new OfficeCompoundStream("1Table", table)
            });
        }
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body", false, 0x11A)]
    [InlineData("body", true, 0x11A)]
    [InlineData("cell", false, 0x11A)]
    [InlineData("cell", true, 0x11A)]
    [InlineData("header", false, 0x122)]
    [InlineData("header", true, 0x122)]
    [InlineData("footer", false, 0x122)]
    [InlineData("footer", true, 0x122)]
    [InlineData("footnote", false, 0x12A)]
    [InlineData("footnote", true, 0x12A)]
    [InlineData("endnote", false, 0x21A)]
    [InlineData("endnote", true, 0x21A)]
    public void LegacyDoc_SectionPagesPreservesEditableInstructionResultAndStoryFormatting(string story, bool complex, int fibOffset) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = AddSectionPagesStoryParagraph(source, story);
        const string instruction = " SECTIONPAGES \\* Roman \\* MERGEFORMAT ";
        Run result = new(new RunProperties(new Bold(), new Color { Val = "225588" }), new Text("II"));
        if (complex) {
            paragraph._paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), result,
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        } else {
            paragraph._paragraph.Append(new SimpleField(result) { Instruction = instruction });
        }

        byte[] native = source.ToBytes(WordFileFormat.Doc);
        Assert.True(OfficeCompoundFileReader.TryRead(native, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = compound!.Streams["WordDocument"];
        byte[] table = compound.Streams["1Table"];
        int offset = BitConverter.ToInt32(word, fibOffset);
        Assert.Equal(22, BitConverter.ToInt32(word, fibOffset + 4));
        Assert.Equal(0x42, table[offset + 17]); // MS-DOC fltSectionPages.
        Assert.Equal(0x80, table[offset + 21]); // The cached result separator is present.

        using WordDocument imported = WordDocument.Load(new MemoryStream(native));
        Assert.Empty(imported.LegacyDocUnsupportedFeatures);
        SimpleField field = Assert.Single(GetSectionPagesStoryRoot(imported, story).Descendants<SimpleField>());
        Assert.Equal(instruction, field.Instruction!.Value);
        Assert.Equal("II", field.InnerText);
        Run display = Assert.Single(field.Elements<Run>());
        Assert.NotNull(display.RunProperties!.Bold);
        Assert.Equal("225588", display.RunProperties.Color!.Val!.Value);
        WordFieldInfo editable = Assert.Single(imported.InspectFields());
        Assert.Equal(WordFieldType.SectionPages, editable.FieldType);
        Assert.Contains(WordFieldFormat.Roman, editable.FormatSwitches);
        Assert.Equal("II", editable.ResultText);

        field.Instruction = " SECTIONPAGES \\* Arabic ";
        display.GetFirstChild<Text>()!.Text = "3";
        using WordDocument edited = WordDocument.Load(new MemoryStream(imported.ToBytes(WordFileFormat.Doc)));
        SimpleField editedField = Assert.Single(GetSectionPagesStoryRoot(edited, story).Descendants<SimpleField>());
        Assert.Equal(" SECTIONPAGES \\* Arabic ", editedField.Instruction!.Value);
        Assert.Equal("3", editedField.InnerText);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_SectionPagesTypedAuthoringRemainsGettable(bool advanced) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph().AddField(WordFieldType.SectionPages, WordFieldFormat.Roman, advanced: advanced);
        using WordDocument imported = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordField field = Assert.Single(imported.Paragraphs, p => p.Field != null).Field!;
        Assert.Equal(WordFieldType.SectionPages, field.FieldType);
        Assert.Contains(WordFieldFormat.Roman, field.FieldFormat);
    }

    [Theory]
    [InlineData("header-table-start5", false, 1)]
    [InlineData("header-table-continuing", false, 2)]
    [InlineData("footer-table-start5", true, 1)]
    [InlineData("footer-table-continuing", true, 2)]
    public void LegacyDoc_SectionPagesImportsWordProducedEditableFields(string fixture, bool footer, int count) {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Word", "LegacyFields", fixture + ".doc");
        using WordDocument document = WordDocument.Load(path);
        MainDocumentPart main = document._wordprocessingDocument!.MainDocumentPart!;
        IEnumerable<SimpleField> fields = footer
            ? main.FooterParts.SelectMany(part => part.Footer.Descendants<SimpleField>())
            : main.HeaderParts.SelectMany(part => part.Header.Descendants<SimpleField>());
        SimpleField[] sectionFields = fields.Where(field => field.Instruction?.Value?.TrimStart().StartsWith("SECTIONPAGES", StringComparison.OrdinalIgnoreCase) == true).ToArray();
        Assert.Equal(count, sectionFields.Length);
        Assert.All(sectionFields, field => Assert.Equal("1", field.InnerText)); // Preserve Word's cached value during import.
        WordFieldInfo[] editable = document.InspectFields().Where(field => field.FieldType == WordFieldType.SectionPages).ToArray();
        Assert.Equal(count, editable.Length);
        Assert.All(editable, field => {
            Assert.Equal("1", field.ResultText);
            Assert.Equal(footer ? WordFieldLocationKind.Footer : WordFieldLocationKind.Header, field.LocationKind);
            Assert.True(field.IsInTable);
        });
    }

    [Theory]
    [InlineData("\u0013 SECTIONPAGES \\* Roman \u0015", "")]
    [InlineData("\u0013 SECTIONPAGES \\* Roman \u0014\u0015", "")]
    [InlineData("\u0013 SECTIONPAGES \\* Roman \u0014II\u0015", "II")]
    public void LegacyDoc_SectionPagesReaderRetainsValidUncachedAndCachedInstructions(string nativeText, string result) {
        LegacyDocTextCharacter[] characters = nativeText.Select((value, index) => new LegacyDocTextCharacter(value, index, index)).ToArray();
        Assert.True(LegacyDocField.TryReadSectionPages(characters, 0, out string instruction, out int start, out int end, out int fieldEnd));
        Assert.Equal(" SECTIONPAGES \\* Roman ", instruction);
        Assert.Equal(result, nativeText.Substring(start, end - start));
        Assert.Equal(nativeText.Length - 1, fieldEnd);
    }

    [Theory]
    [InlineData("\u0013 SECTIONPAGESX \u0014II\u0015")]
    [InlineData("\u0013 SECTIONPAGES \u0014II")]
    [InlineData("\u0013 SECTIONPAGES \r\u0015")]
    public void LegacyDoc_SectionPagesReaderRejectsUnrecognizedOrUnterminatedFields(string nativeText) {
        LegacyDocTextCharacter[] characters = nativeText.Select((value, index) => new LegacyDocTextCharacter(value, index, index)).ToArray();
        Assert.False(LegacyDocField.TryReadSectionPages(characters, 0, out _, out _, out _, out _));
    }

    private static WordParagraph AddSectionPagesStoryParagraph(WordDocument document, string story) {
        if (story == "cell") return document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        if (story == "header" || story == "footer") {
            document.AddHeadersAndFooters();
            return story == "header" ? document.Sections[0].Header.Default!.AddParagraph() : document.Sections[0].Footer.Default!.AddParagraph();
        }
        if (story == "footnote") return document.AddParagraph("Footnote ").AddFootNote("Note ").FootNote!.Paragraphs![1];
        if (story == "endnote") return document.AddParagraph("Endnote ").AddEndNote("Note ").EndNote!.Paragraphs![1];
        return document.AddParagraph();
    }

    private static OpenXmlElement GetSectionPagesStoryRoot(WordDocument document, string story) {
        MainDocumentPart main = document._wordprocessingDocument!.MainDocumentPart!;
        return story switch {
            "header" => main.HeaderParts.Single().Header!,
            "footer" => main.FooterParts.Single().Footer!,
            "footnote" => main.FootnotesPart!.Footnotes!,
            "endnote" => main.EndnotesPart!.Endnotes!,
            _ => main.Document.Body!
        };
    }
}

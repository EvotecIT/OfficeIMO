using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherFormattedEmptyRefillWritingTests {
    public static IEnumerable<object[]> RefillCases() {
        foreach (bool cells in new[] { false, true })
            foreach (string operation in new[] { "interior", "final", "duplicate", "equivalent", "layout" }) yield return new object[] { cells, operation };
    }

    [Theory]
    [MemberData(nameof(RefillCases))]
    public void Count_changing_empty_refill_applies_new_run_styles_and_retains_native_typing_residue(bool cells, string operation) {
        byte[] source = File.ReadAllBytes(Fixture(cells));
        PublisherDocument original = PublisherDocument.Load(source), document = PublisherDocument.Load(source);
        if (operation == "layout") {
            // Leave one genuine empty source among nonuniform nonempty paragraphs.
            if (cells) {
                PublisherTable t = document.Pages[0].Tables.Single();
                document.ReplaceTableCellParagraphs(t.Id, 0, 0, t.Cells[0].Paragraphs.Take(3).ToArray());
            } else document.ReplaceStoryParagraphs(document.TextStories[0].Id, document.TextStories[0].Paragraphs.Take(3).ToArray());
            source = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            original = PublisherDocument.Load(source); document = PublisherDocument.Load(source);
        }
        int page = operation == "equivalent" ? 8 : 0, selected = page == 8 ? 0 : 1;
        PublisherTable? table = cells ? document.Pages[page].Tables.Single() : null;
        PublisherTextStory story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[page];
        var before = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        OfficeRichTextParagraph Filled(int count) => Copy(before[selected], Enumerable.Range(0, count).Select(i => Run("Refill" + i + " ", i)).ToArray());
        OfficeRichTextParagraph first = Filled(operation == "final" ? 1 : 2);
        if (operation == "layout") first = new OfficeRichTextParagraph(first.Runs, OfficeTextAlignment.Justify,
            first.LineHeight, first.Margins, first.Indent, first.LineHeightFactor).WithTabStops(first.TabStops!.WithParagraphAlignment(false));
        OfficeRichTextParagraph[] requested = operation switch {
            "final" => new[] { first }, "duplicate" => new[] { first, before[2], Filled(3) },
            "equivalent" or "layout" => new[] { first, before[2] }, _ => new[] { first, before[2], before[3] }
        };
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(story.Id, requested);
        byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument saved = PublisherDocument.Load(output);
        var actual = cells ? saved.Pages[page].Tables.Single().Cells[0].Paragraphs : saved.TextStories.Single(s => s.Id == story.Id).Paragraphs;
        Assert.Equal(requested.Length, actual.Count);
        for (int i = 0; i < requested.Length; i++) Assert.True(PublisherTextComparison.ParagraphEqual(requested[i], actual[i]), "Requested paragraph " + i);
        byte[] residue = Residue(Definition(original, source, story.Id, selected, "FDPC", true));
        foreach (int paragraph in operation == "duplicate" ? new[] { 0, 2 } : new[] { 0 }) {
            if (operation != "layout") Assert.Equal(Definition(original, source, story.Id, selected, "FDPP"), Definition(saved, output, story.Id, paragraph, "FDPP"));
            int cursor = 0;
            foreach (OfficeRichTextRun run in actual[paragraph].Runs) {
                byte[] body = Definition(saved, output, story.Id, paragraph, "FDPC", position: cursor);
                Assert.Equal(residue, Residue(body)); cursor += run.Text.Length;
            }
            Assert.Equal(Definition(saved, output, story.Id, paragraph, "FDPC", position: cursor - 1), Definition(saved, output, story.Id, paragraph, "FDPC", true));
            Assert.All(actual[paragraph].Runs, r => { Assert.Equal(1033U, r.LanguageId); Assert.Equal(6D, r.AutomaticPairKerningThreshold); });
        }
        foreach (PublisherTextStory neighbor in original.TextStories.Where(s => s.Id != story.Id))
            Assert.True(neighbor.Paragraphs.SequenceEqual(saved.TextStories.Single(s => s.Id == neighbor.Id).Paragraphs, PublisherTextComparison.Paragraphs));
        if (cells) foreach (PublisherTableCell cell in table!.Cells.Skip(1))
            Assert.True(cell.Paragraphs.SequenceEqual(saved.Pages[page].Tables.Single().Cells[cell.RowIndex * table.ColumnCount + cell.ColumnIndex].Paragraphs, PublisherTextComparison.Paragraphs));
        PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS");
        Assert.Equal(output, saved.ToPubBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Formatted_refill_refuses_ambiguous_empty_language_or_opaque_companions(bool opaque) {
        string fixture = Fixture(false); byte[] source = File.ReadAllBytes(fixture);
        if (opaque) {
            PublisherDocument original = PublisherDocument.Load(source);
            var binding = original.SourceStyleCatalog.Stories.Single(s => s.StoryId == original.TextStories[7].Id);
            int cr = binding.Uses.Where(u => u.Tag == "FDPP").ElementAt(1).NativeStart;
            var span = binding.Uses.Single(u => u.Tag == "FDPC" && u.NativeStart <= cr && u.NativeStart + u.NativeLength > cr).DirectDefinition;
            source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
                var fields = new PublisherBlockReader(new PublisherBinaryData(bytes, "test/Quill"), new PublisherParseContext(new PublisherReadOptions(), default)).Chunk(span.Offset, span.Offset + span.Length);
                // This source CR is a separate one-field language override;
                // equivalent empty CRs on page 8 share one native definition.
                PublisherBlock language = Assert.Single(fields, f => f.Id == 0x12); Assert.Equal(0x22, language.Type); Assert.Equal(1036U, language.Value);
                // Preserve its scalar and shape as the native opaque companion.
                bytes[language.Offset - 2] = 0x3E;
            }, "NativeComparison/TextEmptyRefill/native-textemptyrefill.pub");
        }
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory story = document.TextStories[7];
        var runs = new[] { Run("Changed ", 0), Run("typing", 1) };
        if (!opaque) runs = runs.Select(r => r.WithLanguageId(null)).ToArray();
        document.ReplaceStoryParagraphs(story.Id, new[] { Copy(story.Paragraphs[0], runs), story.Paragraphs[2] });
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    private static OfficeRichTextRun Run(string text, int i) => new OfficeRichTextRun(text, i % 2 == 0 ? 12 : 14,
        i % 2 == 0 ? OfficeColor.Green : OfficeColor.Blue, bold: i % 2 != 0, italic: i % 2 == 0,
        fontFamily: i % 2 == 0 ? "Arial" : "Times New Roman", strikethrough: i % 2 == 0,
        underlineStyle: i % 2 == 0 ? OfficeTextDecorationStyle.None : OfficeTextDecorationStyle.Single).WithLanguageId(1033).WithAutomaticPairKerningThreshold(6);

    private static OfficeRichTextParagraph Copy(OfficeRichTextParagraph p, OfficeRichTextRun[] runs) =>
        new OfficeRichTextParagraph(runs, p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);
    private static string Fixture(bool cells) => PublisherNativeTests.Fixture("NativeComparison/" + (cells ? "TableEmptyRefill/native-table" : "TextEmptyRefill/native-text") + "emptyrefill.pub");
    private static byte[] Definition(PublisherDocument d, byte[] source, uint id, int paragraph, string tag, bool cr = false, int position = 0) {
        var binding = d.SourceStyleCatalog.Stories.Single(s => s.StoryId == id);
        var pp = binding.Uses.Where(u => u.Tag == "FDPP").ElementAt(paragraph);
        int native = pp.NativeStart + (cr ? pp.NativeLength - 1 : position);
        var span = tag == "FDPP" ? pp.DirectDefinition : binding.Uses.Single(u => u.Tag == "FDPC" && u.NativeStart <= native && u.NativeStart + u.NativeLength > native).DirectDefinition;
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        return compound!.Streams["Quill/QuillSub/CONTENTS"].Skip(span.Offset).Take(span.Length).ToArray();
    }
    private static byte[] Residue(byte[] bytes) {
        byte[] changed = { 2, 3, 0x0C, 0x0F, 0x10, 0x12, 0x18, 0x1E, 0x24, 0x44, 0x58 };
        var fields = new PublisherBlockReader(new PublisherBinaryData(bytes, "test/residue"), new PublisherParseContext(new PublisherReadOptions(), default)).Chunk(0, bytes.Length);
        return fields.Where(f => !changed.Contains(f.Id)).SelectMany(f => bytes.Skip(f.Offset - 2).Take(f.Length + 2)).ToArray();
    }
}

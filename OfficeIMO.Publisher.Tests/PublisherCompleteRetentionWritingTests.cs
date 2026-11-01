using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherCompleteRetentionWritingTests {
    public static IEnumerable<object[]> Cases() {
        foreach (bool cells in new[] { false, true })
            foreach (string operation in new[] { "drop", "only", "duplicate", "final-source", "moved-final", "equivalent" }) yield return new object[] { cells, operation };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void Complete_nonempty_replacements_preserve_paragraph_and_destination_CR_roles(bool cells, string operation) {
        byte[] source = File.ReadAllBytes(Fixture(cells)); var original = PublisherDocument.Load(source); var d = PublisherDocument.Load(source);
        int page = operation == "equivalent" ? 12 : 0, selected = operation is "final-source" or "moved-final" ? 3 : 0;
        PublisherTable? table = cells ? d.Pages[page].Tables.Single() : null;
        PublisherTextStory story = cells ? d.TextStories.Single(s => s.Id == table!.StoryId) : d.TextStories[page];
        var before = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        OfficeRichTextParagraph Filled(int count) => Copy(before[selected], Enumerable.Range(0, count).Select(i => Run("Rewrite" + i + " ", i)).ToArray());
        OfficeRichTextParagraph[] requested = operation switch {
            "only" => new[] { Filled(4) }, "duplicate" => new[] { Filled(2), before[1], Filled(4), before[2], before[3] },
            "final-source" => new[] { before[2], Filled(1) }, "moved-final" => new[] { Filled(2), before[0], before[1] },
            "equivalent" => new[] { Filled(2), before[2] }, _ => new[] { Filled(2), before[2], before[3] }
        };
        if (cells) d.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested); else d.ReplaceStoryParagraphs(story.Id, requested);
        byte[] output = d.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }); var saved = PublisherDocument.Load(output);
        var actual = cells ? saved.Pages[page].Tables.Single().Cells[0].Paragraphs : saved.TextStories.Single(s => s.Id == story.Id).Paragraphs;
        Assert.Equal(requested.Length, actual.Count);
        for (int p = 0; p < requested.Length; p++) Assert.True(PublisherTextComparison.ParagraphEqual(requested[p], actual[p]), "Requested paragraph " + p);
        var edited = operation == "duplicate" ? new[] { 0, 2 } : operation == "final-source" ? new[] { 1 } : new[] { 0 };
        foreach (int p in edited) {
            Assert.Equal(Definition(original, source, story.Id, selected, "FDPP"), Definition(saved, output, story.Id, p, "FDPP"));
            int crSource = p + 1 == actual.Count ? before.Count - 1 : selected == 3 ? 0 : selected;
            Assert.Equal(CRResidue(Definition(original, source, story.Id, crSource, "FDPC", true)), CRResidue(Definition(saved, output, story.Id, p, "FDPC", true)));
            Assert.All(actual[p].Runs, r => { Assert.Equal(1033U, r.LanguageId); Assert.Equal(6D, r.AutomaticPairKerningThreshold); });
        }
        foreach (PublisherTextStory neighbor in original.TextStories.Where(s => s.Id != story.Id)) {
            Assert.Equal(neighbor.Text, saved.TextStories.Single(s => s.Id == neighbor.Id).Text);
            foreach (string tag in new[] { "FDPC", "FDPP" }) Assert.Equal(NativeDefinitions(original, source, neighbor.Id, tag), NativeDefinitions(saved, output, neighbor.Id, tag));
        }
        if (cells) foreach (PublisherTableCell cell in table!.Cells.Skip(1))
            Assert.True(cell.Paragraphs.SequenceEqual(saved.Pages[page].Tables.Single().Cells[cell.RowIndex * table.ColumnCount + cell.ColumnIndex].Paragraphs, PublisherTextComparison.Paragraphs));
        PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS"); Assert.Equal(output, saved.ToPubBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Different_interior_CR_templates_are_ambiguous_only_for_interior_destinations(bool final) {
        PublisherDocument d = Prepared(false); PublisherTextStory story = d.TextStories[12]; var before = story.Paragraphs;
        OfficeRichTextParagraph replacement = Copy(before[0], new[] { Run("Changed ", 0), Run("body", 1) });
        d.ReplaceStoryParagraphs(story.Id, final ? new[] { replacement } : new[] { replacement, before[2] });
        var options = new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow };
        if (!final) Assert.Throws<OfficeConversionException>(() => d.ToPubBytes(options));
        else {
            PublisherDocument saved = PublisherDocument.Load(d.ToPubBytes(options));
            Assert.True(PublisherTextComparison.ParagraphEqual(replacement, Assert.Single(saved.TextStories.Single(s => s.Id == story.Id).Paragraphs)));
        }
    }

    [Fact]
    public void Different_uniform_candidate_languages_cannot_be_assigned_to_new_text_when_unspecified() {
        PublisherDocument d = Prepared(true); PublisherTextStory story = d.TextStories[12];
        var runs = new[] { Run("Changed ", 0).WithLanguageId(null), Run("body", 1).WithLanguageId(null) };
        d.ReplaceStoryParagraphs(story.Id, new[] { Copy(story.Paragraphs[0], runs), story.Paragraphs[2] });
        Assert.Throws<OfficeConversionException>(() => d.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    private static PublisherDocument Prepared(bool language) {
        var d = PublisherDocument.Load(Fixture(false)); var story = d.TextStories[12]; var p = story.Paragraphs.ToArray(); var runs = p[1].Runs.ToArray();
        for (int i = 0; i < runs.Length; i++) {
            var r = runs[i];
            if (language) runs[i] = r.WithLanguageId(1036);
            else if (i + 1 == runs.Length) runs[i] = new OfficeRichTextRun(r.Text, r.FontSize, r.Color, bold: true, italic: r.Italic,
                fontFamily: r.FontFamily, strikethrough: r.Strikethrough, underlineStyle: r.UnderlineStyle).WithLanguageId(r.LanguageId).WithAutomaticPairKerningThreshold(r.AutomaticPairKerningThreshold);
        }
        p[1] = Copy(p[1], runs); d.ReplaceStoryParagraphs(story.Id, p);
        return PublisherDocument.Load(d.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    private static OfficeRichTextRun Run(string text, int i) => new OfficeRichTextRun(text, i % 2 == 0 ? 12 : 14, i % 2 == 0 ? OfficeColor.Green : OfficeColor.Blue,
        bold: i % 2 != 0, italic: i % 2 == 0, fontFamily: i % 2 == 0 ? "Arial" : "Times New Roman", strikethrough: i % 2 == 0,
        underlineStyle: i % 2 == 0 ? OfficeTextDecorationStyle.None : OfficeTextDecorationStyle.Single).WithLanguageId(1033).WithAutomaticPairKerningThreshold(6);
    private static OfficeRichTextParagraph Copy(OfficeRichTextParagraph p, OfficeRichTextRun[] runs) => new OfficeRichTextParagraph(runs, p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);
    private static string Fixture(bool cells) => PublisherNativeTests.Fixture("NativeComparison/" + (cells ? "TableRunSegmentation/native-table" : "TextRunSegmentation/native-text") + "runsegmentation.pub");
    private static byte[] Definition(PublisherDocument d, byte[] source, uint id, int paragraph, string tag, bool cr = false) {
        var binding = d.SourceStyleCatalog.Stories.Single(s => s.StoryId == id); var pp = binding.Uses.Where(u => u.Tag == "FDPP").ElementAt(paragraph);
        int at = pp.NativeStart + (cr ? pp.NativeLength - 1 : 0); var span = tag == "FDPP" ? pp.DirectDefinition : binding.Uses.Single(u => u.Tag == "FDPC" && u.NativeStart <= at && u.NativeStart + u.NativeLength > at).DirectDefinition;
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        return compound!.Streams["Quill/QuillSub/CONTENTS"].Skip(span.Offset).Take(span.Length).ToArray();
    }
    private static byte[] CRResidue(byte[] bytes) {
        var fields = new PublisherBlockReader(new PublisherBinaryData(bytes, "test/CR"), new PublisherParseContext(new PublisherReadOptions(), default)).Chunk(0, bytes.Length);
        return fields.Where(f => f.Id is not (0x12 or 0x18)).SelectMany(f => bytes.Skip(f.Offset - 2).Take(f.Length + 2)).ToArray();
    }
    private static string[] NativeDefinitions(PublisherDocument d, byte[] source, uint id, string tag) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error); byte[] bytes = compound!.Streams["Quill/QuillSub/CONTENTS"];
        return d.SourceStyleCatalog.Stories.Single(s => s.StoryId == id).Uses.Where(u => u.Tag == tag).SelectMany(u => Enumerable.Repeat(Convert.ToBase64String(bytes, u.DirectDefinition.Offset, u.DirectDefinition.Length), u.NativeLength)).ToArray();
    }
}

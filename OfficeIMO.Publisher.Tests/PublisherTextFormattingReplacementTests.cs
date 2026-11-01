using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTextFormattingReplacementTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Nullable_native_attributes_need_agreement_or_an_explicit_replacement(bool language, bool supplied) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(Fixture(false)));
        PublisherTextStory story = document.TextStories[0];
        var sourceParagraphs = story.Paragraphs.ToArray();
        var sourceRuns = sourceParagraphs[0].Runs.ToArray();
        sourceRuns[1] = language ? sourceRuns[1].WithLanguageId(1036) : sourceRuns[1].WithAutomaticPairKerningThreshold(8);
        sourceParagraphs[0] = Copy(sourceParagraphs[0], sourceRuns);
        document.ReplaceStoryParagraphs(story.Id, sourceParagraphs);
        document = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        story = document.TextStories.Single(s => s.Id == story.Id);
        if (language) Assert.Equal(1036U, story.Paragraphs[0].Runs[1].LanguageId);
        else Assert.Equal(8D, story.Paragraphs[0].Runs[1].AutomaticPairKerningThreshold);
        var runs = new[] { RequestedRun("New first ", 0), RequestedRun("new last", 1) };
        if (!supplied) runs = runs.Select(r => language ? r.WithLanguageId(null) : r.WithAutomaticPairKerningThreshold(null)).ToArray();
        var requested = story.Paragraphs.ToArray(); requested[0] = Copy(requested[0], runs);
        document.ReplaceStoryParagraphs(story.Id, requested);
        var options = new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow };
        if (!supplied) Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(options));
        else {
            PublisherDocument saved = PublisherDocument.Load(document.ToPubBytes(options));
            OfficeRichTextParagraph actual = saved.TextStories.Single(s => s.Id == story.Id).Paragraphs[0];
            Assert.True(PublisherTextComparison.ParagraphEqual(requested[0], actual));
            Assert.All(actual.Runs, r => { Assert.Equal(1033U, r.LanguageId); Assert.Equal(6D, r.AutomaticPairKerningThreshold); });
        }
    }

    [Theory]
    [InlineData(false, 2)]
    [InlineData(false, 4)]
    [InlineData(true, 2)]
    [InlineData(true, 4)]
    public void New_text_and_run_boundaries_apply_complete_formatting_without_losing_native_residue(bool cells, int count) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture(Fixture(cells)));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable? table = cells ? document.Pages[0].Tables.Single() : null;
        PublisherTextStory story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[0];
        IReadOnlyList<OfficeRichTextParagraph> before = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        Assert.Equal(3, before[0].Runs.Count);
        OfficeRichTextRun[] runs = Enumerable.Range(0, count).Select(i => RequestedRun("Fresh" + i + " ", i)).ToArray();
        OfficeRichTextParagraph replacement = Copy(before[0], runs);
        var requested = before.ToArray(); requested[0] = replacement;
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(story.Id, requested);
        byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument saved = PublisherDocument.Load(output);
        var actual = cells ? saved.Pages[0].Tables.Single().Cells[0].Paragraphs : saved.TextStories.Single(s => s.Id == story.Id).Paragraphs;
        Assert.True(PublisherTextComparison.ParagraphEqual(replacement, actual[0]));
        Assert.Equal("Fresh0 Fresh1 " + (count == 4 ? "Fresh2 Fresh3 " : ""), string.Concat(actual[0].Runs.Select(r => r.Text)));
        Assert.All(actual[0].Runs, r => { Assert.Equal(1033U, r.LanguageId); Assert.Equal(6D, r.AutomaticPairKerningThreshold); });
        Assert.Equal(before.Skip(1).Select(p => string.Concat(p.Runs.Select(r => r.Text))), actual.Skip(1).Select(p => string.Concat(p.Runs.Select(r => r.Text))));
        foreach (PublisherTextStory neighbor in document.TextStories.Where(s => s.Id != story.Id)) {
            PublisherTextStory recovered = saved.TextStories.Single(s => s.Id == neighbor.Id);
            Assert.Equal(neighbor.Text, recovered.Text);
            // Formatting pages can coalesce identical adjacent source ranges.
            // Preserve every native definition at each character, including CRs.
            foreach (string tag in new[] { "FDPC", "FDPP" })
                Assert.Equal(NativeDefinitions(document, source, neighbor.Id, tag), NativeDefinitions(saved, output, neighbor.Id, tag));
        }
        if (cells) for (int cell = 1; cell < table!.Cells.Count; cell++)
            Assert.Equal(table.Cells[cell].Text, saved.Pages[0].Tables.Single().Cells[cell].Text);
        byte[][] oldDefinitions = BodyDefinitions(PublisherDocument.Load(source), source, story.Id, before[0]);
        byte[][] newDefinitions = BodyDefinitions(saved, output, story.Id, actual[0]);
        byte[] residue = Residue(oldDefinitions[0]);
        Assert.All(oldDefinitions, definition => Assert.Equal(residue, Residue(definition)));
        Assert.All(newDefinitions, definition => Assert.Equal(residue, Residue(definition)));
        PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS");
        Assert.Equal(output, saved.ToPubBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void New_text_refuses_conflicting_opaque_body_fields_even_with_explicit_language(bool cells) {
        PublisherDocument original = PublisherDocument.Load(PublisherNativeTests.Fixture(Fixture(cells)));
        uint storyId = cells ? original.Pages[0].Tables.Single().StoryId!.Value : original.TextStories[0].Id;
        PublisherQuillStoryStyles binding = original.SourceStyleCatalog.Stories.Single(s => s.StoryId == storyId);
        PublisherQuillSourceSpan span = binding.Uses.Where(u => u.Tag == "FDPC").Skip(1).First().DirectDefinition;
        byte[] source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            var fields = new PublisherBlockReader(new PublisherBinaryData(bytes, "test/Quill"),
                new PublisherParseContext(new PublisherReadOptions(), default)).Chunk(span.Offset, span.Offset + span.Length);
            PublisherBlock opaque = Assert.Single(fields, f => f.Id == 0x3E);
            Assert.Equal(0x22, opaque.Type); Assert.Equal(1033U, opaque.Value);
            PublisherInputContractTests.WriteUInt32(bytes, opaque.Offset, 1036);
        }, Fixture(cells));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable? table = cells ? document.Pages[0].Tables.Single() : null;
        var before = cells ? table!.Cells[0].Paragraphs : document.TextStories.Single(s => s.Id == storyId).Paragraphs;
        var requested = before.ToArray(); requested[0] = Copy(before[0], new[] { RequestedRun("Completely changed ", 0), RequestedRun("body", 1) });
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(storyId, requested);
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    private static OfficeRichTextRun RequestedRun(string text, int index) => new OfficeRichTextRun(text, index % 2 == 0 ? 12 : 14,
        index % 2 == 0 ? OfficeColor.Green : OfficeColor.Blue, bold: index % 2 != 0, italic: index % 2 == 0,
        fontFamily: index % 2 == 0 ? "Arial" : "Times New Roman", strikethrough: index % 2 == 0,
        underlineStyle: index % 2 == 0 ? OfficeTextDecorationStyle.None : OfficeTextDecorationStyle.Single)
        .WithLanguageId(1033).WithAutomaticPairKerningThreshold(6);

    private static OfficeRichTextParagraph Copy(OfficeRichTextParagraph p, IReadOnlyList<OfficeRichTextRun> runs) =>
        new OfficeRichTextParagraph(runs, p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);

    private static byte[][] BodyDefinitions(PublisherDocument document, byte[] source, uint storyId, OfficeRichTextParagraph first) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        byte[] bytes = compound!.Streams["Quill/QuillSub/CONTENTS"];
        PublisherQuillStoryStyles binding = document.SourceStyleCatalog.Stories.Single(s => s.StoryId == storyId);
        int end = binding.NativeStart + first.Runs.Sum(r => r.Text.Length);
        return binding.Uses.Where(u => u.Tag == "FDPC" && u.NativeStart < end && u.NativeStart + u.NativeLength > binding.NativeStart)
            .Select(u => bytes.Skip(u.DirectDefinition.Offset).Take(u.DirectDefinition.Length).ToArray()).ToArray();
    }

    private static byte[] Residue(byte[] bytes) {
        byte[] changed = { 2, 3, 0x0C, 0x0F, 0x10, 0x12, 0x18, 0x1E, 0x24, 0x44, 0x58 };
        var fields = new PublisherBlockReader(new PublisherBinaryData(bytes, "test/native-residue"),
            new PublisherParseContext(new PublisherReadOptions(), default)).Chunk(0, bytes.Length);
        return fields.Where(f => !changed.Contains(f.Id)).SelectMany(f => bytes.Skip(f.Offset - 2).Take(f.Length + 2)).ToArray();
    }

    private static string[] NativeDefinitions(PublisherDocument document, byte[] source, uint storyId, string tag) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        byte[] bytes = compound!.Streams["Quill/QuillSub/CONTENTS"];
        return document.SourceStyleCatalog.Stories.Single(s => s.StoryId == storyId).Uses.Where(u => u.Tag == tag)
            .SelectMany(u => Enumerable.Repeat(Convert.ToBase64String(bytes, u.DirectDefinition.Offset, u.DirectDefinition.Length), u.NativeLength)).ToArray();
    }

    private static string Fixture(bool cells) => "NativeComparison/" + (cells ? "TableRunSegmentation/native-table" : "TextRunSegmentation/native-text") + "runsegmentation.pub";
}

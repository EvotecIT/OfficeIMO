using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherLanguageWritingTests {
    [Fact]
    public void Independent_native_French_cell_language_reaches_the_public_run() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TableCellTyping/native-tablecelltyping.pub"));
        OfficeRichTextRun[] runs = document.TextStories.SelectMany(story => story.Paragraphs).SelectMany(paragraph => paragraph.Runs).ToArray();
        Assert.Contains(runs, run => run.LanguageId == 1036 && run.Text == "Item");
        Assert.Contains(runs, run => run.LanguageId == 1036 && run.Text == "Refilled");
        Assert.Contains(runs, run => run.LanguageId == 1033);
    }

    [Fact]
    public void Authored_language_survives_linked_story_projection_and_overflow() {
        PublisherDocument document = PublisherDocument.Create(300, 400);
        PublisherTextFrame first = document.AddTextFrame(document.Pages[0].Id, 20, 20, 120, 60, new[] {
            new OfficeRichTextParagraph(new[] {
                new OfficeRichTextRun(string.Concat(Enumerable.Repeat("Complete language metadata. ", 40)), 12,
                    OfficeColor.Black, fontFamily: "Arial").WithLanguageId(1033)
            }, lineHeight: 18)
        });
        PublisherPage page = document.AddPage();
        document.AddLinkedTextFrame(first.Id, page.Id, 20, 20, 120, 60);
        Assert.True(document.Pages[1].TextFrames.Single().HasOverflow);
        Assert.All(Assert.Single(document.TextStories).Paragraphs.SelectMany(paragraph => paragraph.Runs), run => Assert.Equal(1033U, run.LanguageId));
        OfficeDrawingRichText[] painted = document.Pages.SelectMany(item => PublisherNativeTests.Elements(item.CreateDrawing())).OfType<OfficeDrawingRichText>().ToArray();
        Assert.Equal(2, painted.Length);
        Assert.All(painted.SelectMany(item => item.Paragraphs).SelectMany(paragraph => paragraph.Runs), run => Assert.Equal(1033U, run.LanguageId));
    }

    [Theory]
    [InlineData("FrameCreation/native-framecreation.pub")]
    [InlineData("GeneratedFrameLayout/generated-frame-layout.pub")]
    [InlineData("GeneratedFrameLayout/native-normalized.pub")]
    public void Language_set_reset_and_unspecified_saves_preserve_opaque_fields_and_neighboring_stories(string fixture) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        uint id = document.TextStories.First().Id;
        var identities = document.Pages.SelectMany(page => page.TextFrames).Select(frame => (frame.Id, frame.StoryId, frame.PageId)).ToArray();
        var stories = document.TextStories.Select(story => (story.Id, story.Text)).ToArray();
        foreach (uint? language in new uint?[] { 1036, null, 1033 }) {
            PublisherTextStory story = document.TextStories.Single(value => value.Id == id);
            document.ReplaceStoryParagraphs(id, story.Paragraphs.Select(paragraph => new OfficeRichTextParagraph(
                paragraph.Runs.Select(run => run.WithLanguageId(language)).ToArray(), paragraph.Alignment,
                paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops)).ToArray());
            byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS");
            Assert.Equal(OpaqueCharacterFields(source), OpaqueCharacterFields(output));
            document = PublisherDocument.Load(output);
            Assert.Equal(stories, document.TextStories.Select(value => (value.Id, value.Text)));
            Assert.Equal(identities, document.Pages.SelectMany(page => page.TextFrames).Select(frame => (frame.Id, frame.StoryId, frame.PageId)));
            Assert.All(document.TextStories.Single(value => value.Id == id).Paragraphs.SelectMany(paragraph => paragraph.Runs),
                run => Assert.Equal(language ?? 1036, run.LanguageId));
            source = output;
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Unqualified_language_scalar_remains_opaque_and_refuses_explicit_editing(bool outOfRange) {
        byte[] source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            var data = new PublisherBinaryData(bytes, "Quill");
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            PublisherQuillChunk page = PublisherQuillDirectory.Read(data, context).First(chunk => chunk.Tag == "FDPC");
            int count = data.U16(page.Offset);
            int position = page.Offset + data.U16(page.Offset + 8 + count * 4);
            PublisherBlock language = new PublisherBlockReader(data, context).Chunk(position, page.End).Single(field => field.Id == 0x12);
            if (outOfRange) PublisherInputContractTests.WriteUInt32(bytes, language.Offset, 65536);
            else bytes[language.Offset - 1] = 0x20;
        }, "NativeComparison/FrameCreation/native-framecreation.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_TEXT_LANGUAGE_UNASSESSED");
        Assert.Equal(source, document.ToPubBytes());
        PublisherTextStory story = document.TextStories.First();
        document.ReplaceStoryParagraphs(story.Id, story.Paragraphs.Select(paragraph => new OfficeRichTextParagraph(
            paragraph.Runs.Select(run => run.WithLanguageId(1036)).ToArray(), paragraph.Alignment, paragraph.LineHeight,
            paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops)).ToArray());
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Retained_count_changes_apply_language_with_or_without_text_replacement(bool cells, bool replaceText) {
        string leaf = cells ? "TableNonuniformParagraphs/native-tablenonuniformparagraphs.pub"
            : "TextNonuniformParagraphs/native-textnonuniformparagraphs.pub";
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/" + leaf));
        PublisherTable? table = cells ? document.Pages[0].Tables.Single() : null;
        PublisherTextStory story = cells ? document.TextStories.Single(value => value.Id == table!.StoryId) : document.TextStories[0];
        var before = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        var requested = before.Skip(1).Select(paragraph => new OfficeRichTextParagraph(paragraph.Runs.Select(run =>
            new OfficeRichTextRun(replaceText ? run.Text + " extra" : run.Text, run.FontSize, run.Color, run.Bold,
                run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle,
                run.StrikethroughStyle, run.Baseline, run.ParagraphIndent).WithLanguageId(1036)).ToArray(),
            paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor)
            .WithTabStops(paragraph.TabStops)).ToArray();
        Assert.NotEmpty(requested);
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(story.Id, requested);
        PublisherDocument saved = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        var actual = cells ? saved.Pages[0].Tables.Single().Cells[0].Paragraphs : saved.TextStories[0].Paragraphs;
        Assert.Equal(requested.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))),
            actual.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));
        Assert.All(actual.SelectMany(paragraph => paragraph.Runs), run => Assert.Equal(1036U, run.LanguageId));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Empty_typing_topology_applies_explicit_language_to_each_retained_CR(bool cells) {
        string leaf = cells ? "TableCellTyping/native-tablecelltyping.pub" : "TextUniformEmptyParagraphs/native-textuniformemptyparagraphs.pub";
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/" + leaf));
        if (!cells) {
            PublisherTextStory emptySource = document.TextStories[0];
            document.ReplaceStoryParagraphs(emptySource.Id, new[] { emptySource.Paragraphs[0] });
            document = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        }
        PublisherTable? table = cells ? document.Pages.SelectMany(page => page.Tables).First(item =>
            item.Cells.Any(cell => cell.Paragraphs.Count == 1 && cell.Paragraphs[0].Runs.All(run => run.Text.Length == 0))) : null;
        PublisherTableCell? cell = table?.Cells.First(value => value.Paragraphs.Count == 1 && value.Paragraphs[0].Runs.All(run => run.Text.Length == 0));
        PublisherTextStory story = cells ? document.TextStories.Single(value => value.Id == table!.StoryId)
            : document.TextStories.First(value => value.Paragraphs.Count == 1 && value.Paragraphs[0].Runs.All(run => run.Text.Length == 0));
        OfficeRichTextParagraph original = cells ? cell!.Paragraphs[0] : story.Paragraphs.Single();
        var typed = new OfficeRichTextParagraph(original.Runs.Select(run => run.WithLanguageId(1036)).ToArray(),
            original.Alignment, original.LineHeight, original.Margins, original.Indent, original.LineHeightFactor).WithTabStops(original.TabStops);
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, cell!.RowIndex, cell.ColumnIndex, new[] { typed, typed });
        else document.ReplaceStoryParagraphs(story.Id, new[] { typed, typed });
        PublisherDocument saved = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        var actual = cells ? saved.Pages.SelectMany(page => page.Tables).Single(value => value.Id == table!.Id)
            .Cells.Single(value => value.RowIndex == cell!.RowIndex && value.ColumnIndex == cell.ColumnIndex).Paragraphs : saved.TextStories.Single(value => value.Id == story.Id).Paragraphs;
        Assert.Equal(2, actual.Count);
        Assert.All(actual.SelectMany(paragraph => paragraph.Runs), run => { Assert.Equal(string.Empty, run.Text); Assert.Equal(1036U, run.LanguageId); });
    }

    private static string[] OpaqueCharacterFields(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        var data = new PublisherBinaryData(compound!.Streams["Quill/QuillSub/CONTENTS"], "Quill");
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        var blocks = new PublisherBlockReader(data, context);
        return PublisherQuillDirectory.Read(data, context).Where(chunk => chunk.Tag == "FDPC").SelectMany(chunk => {
            int count = data.U16(chunk.Offset);
            return Enumerable.Range(0, count).Select(index => {
                int position = chunk.Offset + data.U16(chunk.Offset + 8 + count * 4 + index * 2);
                return string.Join("|", blocks.Chunk(position, chunk.End).Where(field => field.Id != 0x12)
                    .Select(field => Convert.ToBase64String(data.Bytes, field.Offset - 2, field.Length + 2)));
            });
        }).ToArray();
    }
}

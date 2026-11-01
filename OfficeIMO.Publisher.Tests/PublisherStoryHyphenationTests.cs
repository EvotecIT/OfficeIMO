using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherStoryHyphenationTests {
    private const string NativeText = "NativeComparison/Text/native-text.pub";
    private const string Quill = "Quill/QuillSub/CONTENTS";
    private const long MaximumBytes = 64L * 1024 * 1024;

    [Theory]
    [InlineData(NativeText, false)]
    [InlineData("NativeComparison/TextHyphenation/native-paragraph-suppressed.pub", true)]
    public void Paragraph_suppression_is_read_and_editable_without_changing_story_text(string fixture, bool initiallySuppressed) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture(fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory story = document.TextStories[0];
        Assert.Equal(true, story.AutomaticHyphenation);
        Assert.All(story.Paragraphs, p => Assert.Equal(initiallySuppressed, p.SuppressAutomaticHyphenation));
        Assert.Equal(source, document.ToPubBytes());
        document.ReplaceStoryParagraphs(story.Id, story.Paragraphs.Select(p =>
            p.WithAutomaticHyphenationSuppressed(!initiallySuppressed).WithTabStops(p.TabStops)).ToArray());
        byte[] saved = Save(document);
        Assert.Equal(Compound(source).Streams["Contents"], Compound(saved).Streams["Contents"]);
        PublisherDocument reloaded = PublisherDocument.Load(saved);
        PublisherTextStory updated = reloaded.TextStories.Single(s => s.Id == story.Id);
        Assert.All(updated.Paragraphs, p => Assert.Equal(!initiallySuppressed, p.SuppressAutomaticHyphenation));
        Assert.Equal(story.Text, updated.Text);
        Assert.Equal(true, updated.AutomaticHyphenation);
        Assert.Equal(document.TextStories.Select(s => (s.Id, s.Text)), reloaded.TextStories.Select(s => (s.Id, s.Text)));
        Assert.All(reloaded.TextStories.Where(s => s.Id != story.Id).SelectMany(s => s.Paragraphs),
            p => Assert.Equal(initiallySuppressed, p.SuppressAutomaticHyphenation));
        Assert.Equal(saved, reloaded.ToPubBytes());
        Assert.All(story.Paragraphs, p => Assert.Equal(initiallySuppressed, p.SuppressAutomaticHyphenation));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Unknown_paragraph_flag_is_preserved_and_refuses_an_explicit_setting_edit(bool suppressed) {
        const string fixture = "NativeComparison/TextHyphenation/native-paragraph-suppressed.pub";
        byte[] source = PublisherInputContractTests.Mutate(Quill, bytes => {
            var context = new PublisherParseContext(new PublisherReadOptions(), default);
            var data = new PublisherBinaryData(bytes, Quill); var blocks = new PublisherBlockReader(data, context);
            var page = PublisherQuillDirectory.Read(data, context).First(chunk => chunk.Tag == "FDPP");
            int offset = page.Offset + data.U16(page.Offset + 8 + data.U16(page.Offset) * 4);
            var flag = PublisherContentsReader.Field(blocks.Chunk(offset, page.End), 0x1C)!.Value;
            bytes[flag.Offset - 1] = 0x08;
        }, fixture);
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory story = document.TextStories[0];
        Assert.Contains(document.ReadReport.FidelityDiagnostics, d => d.Code == "PUB_HYPHENATION_SUPPRESSION_UNASSESSED");
        Assert.Null(story.Paragraphs[0].SuppressAutomaticHyphenation);
        Assert.Equal(source, document.ToPubBytes());
        document.ReplaceStoryParagraphs(story.Id, story.Paragraphs.Select(p => p.WithAutomaticHyphenationSuppressed(suppressed)).ToArray());
        Assert.False(document.PrepareNativeBytes(MaximumBytes).IsPrepared);
        Assert.Throws<OfficeConversionException>(() => Save(document));
    }

    [Theory]
    [InlineData(NativeText)]
    [InlineData("NativeComparison/TextHyphenation/native-texthyphenation.pub")]
    public void Native_story_flag_round_trips_without_changing_text_formatting_or_other_stories(string fixture) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture(fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory original = document.TextStories[0];
        Assert.Equal(true, original.AutomaticHyphenation);
        Assert.Equal(source, document.ToPubBytes());
        document.SetStoryAutomaticHyphenation(original.Id, true);
        Assert.Equal(source, document.ToPubBytes());

        document.SetStoryAutomaticHyphenation(original.Id, false);
        document.ReplaceStoryParagraphs(original.Id, original.Paragraphs);
        PublisherTextStory edited = document.TextStories.Single(s => s.Id == original.Id);
        Assert.Equal(false, edited.AutomaticHyphenation);
        Assert.Equal(true, original.AutomaticHyphenation);
        byte[] disabled = Save(document);
        Assert.Equal(Compound(source).Streams[Quill], Compound(disabled).Streams[Quill]);
        PublisherDocument reloaded = PublisherDocument.Load(disabled);
        Assert.Equal(document.TextStories.Select(s => (s.Id, s.Text, s.AutomaticHyphenation)),
            reloaded.TextStories.Select(s => (s.Id, s.Text, s.AutomaticHyphenation)));
        Assert.Equal(disabled, reloaded.ToPubBytes());
        Assert.All(reloaded.TextStories.Where(s => s.Id != original.Id), s => Assert.Equal(true, s.AutomaticHyphenation));

        reloaded.SetStoryAutomaticHyphenation(original.Id, true);
        byte[] enabled = Save(reloaded);
        Assert.Equal(true, PublisherDocument.Load(enabled).TextStories.Single(s => s.Id == original.Id).AutomaticHyphenation);
        Assert.Equal(Compound(source).Streams[Quill], Compound(enabled).Streams[Quill]);
    }

    [Fact]
    public void Story_flag_composes_with_native_text_replacement_and_contents_relocation() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(NativeText));
        PublisherTextStory story = Assert.Single(document.TextStories);
        OfficeRichTextParagraph original = Assert.Single(story.Paragraphs);
        OfficeRichTextRun run = Assert.Single(story.Paragraphs[0].Runs);
        const string replacement = "publication hyphenation automatic";
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(replacement, run.FontSize, run.Color,
            run.Bold, run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor,
            run.UnderlineStyle, run.StrikethroughStyle, run.Baseline, run.ParagraphIndent) {
                AutomaticPairKerningThreshold = run.AutomaticPairKerningThreshold, LanguageId = run.LanguageId
            } }, original.Alignment, original.LineHeight, original.Margins, original.Indent, original.LineHeightFactor).WithTabStops(original.TabStops);
        document.SetStoryAutomaticHyphenation(story.Id, false);
        document.ReplaceStoryParagraphs(story.Id, new[] { paragraph });
        PublisherTextFrame frame = Assert.Single(document.Pages[0].TextFrames);
        document.SetTextFrameLayout(frame.Id, frame.ColumnCount, frame.ColumnSpacing, frame.Padding, OfficeTextVerticalAlignment.Bottom);
        var prepared = document.PrepareNativeBytes(MaximumBytes);
        Assert.True(prepared.IsPrepared, string.Join("\n", prepared.FidelityDiagnostics.Select(d => d.Code + ": " + d.Message)));
        PublisherDocument reloaded = PublisherDocument.Load(Save(document));
        PublisherTextStory saved = Assert.Single(reloaded.TextStories);
        Assert.Equal(false, saved.AutomaticHyphenation);
        Assert.Equal(replacement + "\n", saved.Text);
        Assert.Equal(OfficeTextVerticalAlignment.Bottom, Assert.Single(reloaded.Pages[0].TextFrames).VerticalAlignment);
    }

    [Fact]
    public void Missing_catalog_remains_unknown_and_cannot_silently_discard_a_setting_edit() {
        byte[] source = PublisherInputContractTests.Mutate("Contents", bytes => {
            var context = new PublisherParseContext(new PublisherReadOptions(), default);
            var data = new PublisherBinaryData(bytes, "Contents"); var blocks = new PublisherBlockReader(data, context);
            var contents = new PublisherContentsReader(data, context).Read();
            var document = contents.Chunks.Values.Single(chunk => chunk.Kind == 0x44);
            var reference = PublisherContentsReader.Field(blocks.Chunk(document.Offset, document.End), 0x20)!.Value;
            bytes[reference.Offset - 2] = 0xFE;
        }, NativeText);
        AssertOpaqueSetting(source, expectedDiagnostic: false);
    }

    [Fact]
    public void Unknown_native_flag_is_preserved_and_cannot_be_overwritten_as_a_qualified_boolean() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(NativeText));
        document.SetStoryAutomaticHyphenation(Assert.Single(document.TextStories).Id, false);
        byte[] source = Save(document);
        OfficeCompoundFile compound = Compound(source);
        byte[] contentsBytes = (byte[])compound.Streams["Contents"].Clone();
        var context = new PublisherParseContext(new PublisherReadOptions(), default);
        var data = new PublisherBinaryData(contentsBytes, "Contents"); var blocks = new PublisherBlockReader(data, context);
        var contents = new PublisherContentsReader(data, context).Read();
        var catalog = contents.Chunks.Values.Single(chunk => chunk.Kind == 0x65);
        var array = PublisherContentsReader.Field(blocks.Chunk(catalog.Offset, catalog.End), 2)!.Value;
        var entry = Assert.Single(blocks.Children(array));
        var flag = PublisherContentsReader.Field(blocks.Children(entry), 4)!.Value;
        contentsBytes[flag.Offset - 1] = 0x05;
        AssertOpaqueSetting(OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Contents"] = contentsBytes }), expectedDiagnostic: true);
    }

    [Fact]
    public void Setting_edit_checks_identity_cancellation_and_table_ownership_without_mutation() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(NativeText));
        PublisherTextStory story = Assert.Single(document.TextStories);
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetStoryAutomaticHyphenation(uint.MaxValue, false));
        Assert.Throws<OperationCanceledException>(() => document.SetStoryAutomaticHyphenation(story.Id, false, new CancellationToken(true)));
        Assert.Same(story, Assert.Single(document.TextStories));
        PublisherDocument table = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TableCellParagraphs/native-tablecellparagraphs.pub"));
        uint tableStory = table.Pages.SelectMany(page => page.Tables).First().StoryId!.Value;
        var snapshots = table.TextStories.ToArray();
        Assert.Throws<NotSupportedException>(() => table.SetStoryAutomaticHyphenation(tableStory, false));
        Assert.Equal(snapshots, table.TextStories);
    }

    private static void AssertOpaqueSetting(byte[] source, bool expectedDiagnostic) {
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory story = Assert.Single(document.TextStories);
        Assert.Null(story.AutomaticHyphenation);
        Assert.Equal(source, document.ToPubBytes());
        Assert.Equal(expectedDiagnostic, document.ReadReport.FidelityDiagnostics.Any(d => d.Code == "PUB_HYPHENATION_SETTING_UNASSESSED"));
        document.SetStoryAutomaticHyphenation(story.Id, false);
        PublisherNativePreparation prepared = document.PrepareNativeBytes(MaximumBytes);
        Assert.False(prepared.IsPrepared);
        Assert.Contains(prepared.FidelityDiagnostics, d => d.LossKind == OfficeConversionLossKind.Failure && d.Location == "Contents/story/" + story.Id);
        Assert.Throws<OfficeConversionException>(() => Save(document));
    }

    private static byte[] Save(PublisherDocument document) => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
    private static OfficeCompoundFile Compound(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, new OfficeCompoundReadOptions(), default, out var compound, out var error), error);
        return compound!;
    }
}

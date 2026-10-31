using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherNativeFrameLayoutTests {
    [Fact]
    public void Implicit_native_column_gap_is_read_and_overridden_when_adding_columns() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/ImplicitColumnGap/implicit-column-gap.pub"));
        PublisherTextFrame first = Assert.Single(document.Pages[0].TextFrames);
        PublisherTextFrame last = Assert.Single(document.Pages[2].TextFrames);
        Assert.Equal(7.2, first.ColumnSpacing, 6);
        Assert.Equal(7.2, last.ColumnSpacing, 6);

        document.SetTextFrameLayout(first.Id, 2, 0, first.Padding, first.VerticalAlignment);
        PublisherDocument reopened = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        PublisherTextFrame actual = Assert.Single(reopened.Pages[0].TextFrames);
        Assert.Equal(2, actual.ColumnCount);
        Assert.Equal(0, actual.ColumnSpacing);
        Assert.Equal(7.2, Assert.Single(reopened.Pages[2].TextFrames).ColumnSpacing, 6);
        Assert.Empty(reopened.Pages[1].TextFrames);
        Assert.Equal(document.TextStories.Select(story => (story.Id, story.Text)), reopened.TextStories.Select(story => (story.Id, story.Text)));
    }

    [Theory]
    [InlineData("Text")]
    [InlineData("TextFrameLayout")]
    [InlineData("MultipleMasters")]
    [InlineData("Newsletter")]
    public void Native_layout_saves_grow_missing_fields_and_preserve_complete_stories_and_object_ownership(string fixture) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture($"NativeComparison/{fixture}/native-{fixture.ToLowerInvariant()}.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextFrame[] frames = document.Pages.Concat(document.MasterPages).SelectMany(page => page.TextFrames)
            .Where(frame => frame.CanEditLayout).ToArray();
        Assert.NotEmpty(frames);
        var expected = frames.Select((frame, index) => new {
            frame.Id, frame.StoryId, frame.PageId, frame.PreviousFrameId, frame.NextFrameId,
            Columns = index % 3 + 1, Spacing = 12.1,
            Padding = new OfficeTextPadding(1.21, 2.42, 3.63, 4.84),
            Alignment = (OfficeTextVerticalAlignment)(index % 3)
        }).ToArray();
        foreach (var value in expected)
            document.SetTextFrameLayout(value.Id, value.Columns, value.Spacing, value.Padding, value.Alignment);
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        byte[] candidate = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument reopened = PublisherDocument.Load(candidate);
        Assert.Equal(document.TextStories.Select(story => (story.Id, story.Text)), reopened.TextStories.Select(story => (story.Id, story.Text)));
        Assert.Equal(document.Pages.Select(page => (page.Id, page.MasterPageId)), reopened.Pages.Select(page => (page.Id, page.MasterPageId)));
        var actual = reopened.Pages.Concat(reopened.MasterPages).SelectMany(page => page.TextFrames).ToDictionary(frame => frame.Id);
        foreach (var value in expected) {
            PublisherTextFrame frame = actual[value.Id];
            Assert.Equal((value.StoryId, value.PageId, value.PreviousFrameId, value.NextFrameId),
                (frame.StoryId, frame.PageId, frame.PreviousFrameId, frame.NextFrameId));
            Assert.Equal(value.Columns, frame.ColumnCount); Assert.Equal(value.Spacing, frame.ColumnSpacing, 8);
            Assert.Equal(value.Padding, frame.Padding); Assert.Equal(value.Alignment, frame.VerticalAlignment);
        }
        OfficeCompoundFile before = Compound(source), after = Compound(candidate);
        Assert.Equal(before.Streams.Keys.OrderBy(key => key), after.Streams.Keys.OrderBy(key => key));
        foreach (string name in before.Streams.Keys.Where(key => key != "Contents" && key != "Escher/EscherStm" && key != "\u0005SummaryInformation"))
            Assert.Equal(before.Streams[name], after.Streams[name]);
        Assert.Equal(source, File.ReadAllBytes(PublisherNativeTests.Fixture($"NativeComparison/{fixture}/native-{fixture.ToLowerInvariant()}.pub")));
        foreach (PublisherTextFrame original in frames)
            reopened.SetTextFrameLayout(original.Id, original.ColumnCount, original.ColumnSpacing, original.Padding, original.VerticalAlignment);
        PublisherDocument reset = PublisherDocument.Load(reopened.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        foreach (PublisherTextFrame original in frames) {
            PublisherTextFrame frame = reset.Pages.Concat(reset.MasterPages).SelectMany(page => page.TextFrames).Single(item => item.Id == original.Id);
            Assert.Equal(original.ColumnCount, frame.ColumnCount); Assert.Equal(original.ColumnSpacing, frame.ColumnSpacing);
            Assert.Equal(original.Padding, frame.Padding); Assert.Equal(original.VerticalAlignment, frame.VerticalAlignment);
        }
    }

    [Fact]
    public void Native_layout_save_refuses_inexact_emu_values_instead_of_rounding_or_copying_source() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Text/native-text.pub"));
        PublisherTextFrame frame = Assert.Single(document.Pages[0].TextFrames);
        document.SetTextFrameLayout(frame.Id, 2, 12, new OfficeTextPadding(0.5 / 12700, 0, 0, 0), OfficeTextVerticalAlignment.Top);
        var preparation = document.PrepareNativeBytes(64 * 1024 * 1024);
        Assert.False(preparation.IsPrepared);
        Assert.Contains(preparation.FidelityDiagnostics, diagnostic => diagnostic.Message.Contains("EMU"));
    }

    private static OfficeCompoundFile Compound(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        return compound!;
    }

    [Theory]
    [InlineData("length")]
    [InlineData("directory-type")]
    public void Growing_alignment_refuses_unqualified_relocation_headers_without_changing_exact_source_copy(string invalid) {
        byte[] source = PublisherInputContractTests.Mutate("Contents", bytes => {
            if (invalid == "length") { bytes[8] ^= 1; return; }
            var data = new PublisherBinaryData(bytes, "Contents");
            var blocks = new PublisherBlockReader(data, new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None));
            PublisherBlock directory = blocks.Chunk(data.Offset(data.U32(0x1A))).Single(field => field.Type == 0x90);
            PublisherBlock slot = blocks.Children(directory).First(field => field.Type == 0x88);
            PublisherBlock offset = blocks.Children(slot).Single(field => field.Id == 4);
            bytes[offset.Offset - 1] = 0x20;
        }, "NativeComparison/Text/native-text.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        PublisherTextFrame frame = Assert.Single(document.Pages[0].TextFrames);
        document.SetTextFrameLayout(frame.Id, frame.ColumnCount, frame.ColumnSpacing, frame.Padding, OfficeTextVerticalAlignment.Center);
        PublisherNativePreparation result = document.PrepareNativeBytes(64 * 1024 * 1024);
        Assert.False(result.IsPrepared);
        Assert.Contains(result.FidelityDiagnostics, diagnostic => diagnostic.Message.Contains("relocation"));
    }
}

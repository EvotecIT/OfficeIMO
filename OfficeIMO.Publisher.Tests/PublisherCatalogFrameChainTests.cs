using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherCatalogFrameChainTests {
    private static readonly string Newsletter = Path.Combine("NativeComparison", "Newsletter", "native-newsletter.pub");

    [Fact]
    public void Unique_shape_ordinals_without_a_native_catalog_do_not_guess_a_story_chain() {
        byte[] input = PublisherInputContractTests.Mutate("Contents", bytes => {
            var data = new PublisherBinaryData(bytes, "Contents");
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            var blocks = new PublisherBlockReader(data, context);
            PublisherContents publication = new PublisherContentsReader(data, context).Read();
            PublisherContentChunk document = publication.Chunks.Values.Single(chunk => chunk.Kind == 0x44);
            PublisherBlock catalog = PublisherContentsReader.Field(blocks.Chunk(document.Offset, document.End), 0x20)!.Value;
            bytes[catalog.Offset - 2] = 0xFE;
        }, Newsletter);

        PublisherDocument document = PublisherDocument.Load(input);
        PublisherTextStory story = Assert.Single(document.TextStories, item => item.Id == 3);
        Assert.Contains("Section 10.", story.Text);
        Assert.DoesNotContain(document.Pages.SelectMany(page => page.TextFrames), frame => frame.StoryId == 3);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_CHAIN_UNRESOLVED");
    }

    [Fact]
    public void Catalog_chain_rejects_a_shape_ordinal_that_disagrees_with_its_binding() {
        byte[] input = PublisherInputContractTests.Mutate("Contents", bytes => {
            int ordinal = PublisherTextFlowTests.NativeFrameField(bytes, 298, 0x28);
            PublisherInputContractTests.WriteUInt32(bytes, ordinal, 2);
        }, Newsletter);

        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input));
    }

    [Theory]
    [InlineData(0x61U, 1U, (byte)2, 0U)] // Duplicate ordinal in the frame binding catalog.
    [InlineData(0x61U, 1U, (byte)3, 296U)] // Binding points to another story's title frame.
    [InlineData(0x65U, 0U, (byte)7, 3U)] // Declared frame count disagrees with the complete mapping.
    public void Catalog_chain_rejects_duplicate_foreign_or_incomplete_bindings(uint kind, uint ordinal, byte field, uint replacement) {
        byte[] input = PublisherInputContractTests.Mutate("Contents", bytes => {
            var data = new PublisherBinaryData(bytes, "Contents");
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            var blocks = new PublisherBlockReader(data, context);
            PublisherContents publication = new PublisherContentsReader(data, context).Read();
            PublisherContentChunk chunk = publication.Chunks.Values.Single(item => item.Kind == kind);
            PublisherBlock array = PublisherContentsReader.Field(blocks.Chunk(chunk.Offset, chunk.End), 2)!.Value;
            IReadOnlyList<PublisherBlock> entry = blocks.Children(array)
                .Select(blocks.Children)
                .Single(fields => PublisherContentsReader.Field(fields, 1)!.Value.Value == 3
                    && (kind != 0x61 || (PublisherContentsReader.Field(fields, 2)?.Value ?? 0) == ordinal));
            PublisherBlock changed = PublisherContentsReader.Field(entry, field)!.Value;
            PublisherInputContractTests.WriteUInt32(bytes, changed.Offset, replacement);
        }, Newsletter);

        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input));
    }
}

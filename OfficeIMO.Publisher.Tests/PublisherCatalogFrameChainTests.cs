using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherCatalogFrameChainTests {
    private static readonly string Newsletter = Path.Combine("NativeComparison", "Newsletter", "native-newsletter.pub");

    [Theory]
    [InlineData(2U, new uint[] { 292, 294 })]
    [InlineData(18U, new uint[] { 326, 327, 328 })]
    public void Native_catalog_chains_with_independent_story_metadata_match_explicit_link_placement(uint storyId, uint[] expectedOrder) {
        const string fixture = "SampleNewsletter.pub";
        PublisherDocument original = PublisherDocument.Load(PublisherNativeTests.Fixture(fixture));
        byte[] input = PublisherInputContractTests.Mutate("Contents", bytes => {
            var data = new PublisherBinaryData(bytes, "Contents");
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            var blocks = new PublisherBlockReader(data, context);
            PublisherContents publication = new PublisherContentsReader(data, context).Read();
            PublisherContentChunk storyCatalog = publication.Chunks.Values.Single(chunk => chunk.Kind == 0x65);
            PublisherBlock array = PublisherContentsReader.Field(blocks.Chunk(storyCatalog.Offset, storyCatalog.End), 2)!.Value;
            IReadOnlyList<PublisherBlock> entry = blocks.Children(array).Select(blocks.Children)
                .Single(fields => PublisherContentsReader.Field(fields, 1)!.Value.Value == storyId);
            PublisherBlock metadata = PublisherContentsReader.Field(entry, 7)!.Value;
            // Native IdLayout metadata is independent of the number of frames.
            // Keep this fixture a counterexample to the former count predicate.
            Assert.NotEqual((uint)expectedOrder.Length, metadata.Value);
            foreach (uint frameId in expectedOrder) {
                PublisherSourceShape frame = publication.Shapes[frameId];
                foreach (PublisherBlock link in frame.Fields.Where(field => field.Id is 0x36 or 0x37)) {
                    byte unusedField = link.Id == 0x36 ? (byte)0xFE : (byte)0xFF;
                    Assert.DoesNotContain(frame.Fields, field => field.Id == unusedField);
                    // Derive an ordinal-only input without changing any native
                    // story metadata, frame catalog entry, anchor or text byte.
                    bytes[link.Offset - 2] = unusedField;
                }
            }
        }, fixture);

        PublisherDocument recovered = PublisherDocument.Load(input);
        PublisherTextStory originalStory = Assert.Single(original.TextStories, story => story.Id == storyId);
        Assert.NotEmpty(originalStory.Text);
        Assert.Equal(originalStory.Text, Assert.Single(recovered.TextStories, story => story.Id == storyId).Text);
        PublisherTextFrame[] originalFrames = Frames(original, storyId);
        PublisherTextFrame[] recoveredFrames = Frames(recovered, storyId);
        Assert.Equal(expectedOrder, originalFrames.Select(frame => frame.Id));
        Assert.Equal(expectedOrder, recoveredFrames.Select(frame => frame.Id));
        for (int i = 0; i < originalFrames.Length; i++) {
            PublisherTextFrame expected = originalFrames[i], actual = recoveredFrames[i];
            Assert.Equal(expected.PageId, actual.PageId);
            Assert.Equal(expected.PreviousFrameId, actual.PreviousFrameId);
            Assert.Equal(expected.NextFrameId, actual.NextFrameId);
            Assert.Equal((expected.X, expected.Y, expected.Width, expected.Height), (actual.X, actual.Y, actual.Width, actual.Height));
            Assert.Equal((expected.TextStart, expected.TextLength, expected.HasOverflow), (actual.TextStart, actual.TextLength, actual.HasOverflow));
        }
    }

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
    [InlineData(0x61U, 3U, 1U, (byte)2, 0U)] // Duplicate ordinal in the frame binding catalog.
    [InlineData(0x61U, 3U, 1U, (byte)3, 296U)] // Binding points to another story's title frame.
    [InlineData(0x61U, 3U, 1U, (byte)1, 2U)] // One binding is assigned elsewhere, leaving an incomplete chain.
    [InlineData(0x65U, 2U, 0U, (byte)1, 3U)] // Two story catalog entries declare the same candidate story.
    public void Catalog_chain_rejects_duplicate_foreign_or_incomplete_bindings(uint kind, uint storyId, uint ordinal, byte field, uint replacement) {
        byte[] input = PublisherInputContractTests.Mutate("Contents", bytes => {
            var data = new PublisherBinaryData(bytes, "Contents");
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            var blocks = new PublisherBlockReader(data, context);
            PublisherContents publication = new PublisherContentsReader(data, context).Read();
            PublisherContentChunk chunk = publication.Chunks.Values.Single(item => item.Kind == kind);
            PublisherBlock array = PublisherContentsReader.Field(blocks.Chunk(chunk.Offset, chunk.End), 2)!.Value;
            IReadOnlyList<PublisherBlock> entry = blocks.Children(array)
                .Select(blocks.Children)
                .Single(fields => PublisherContentsReader.Field(fields, 1)!.Value.Value == storyId
                    && (kind != 0x61 || (PublisherContentsReader.Field(fields, 2)?.Value ?? 0) == ordinal));
            PublisherBlock changed = PublisherContentsReader.Field(entry, field)!.Value;
            PublisherInputContractTests.WriteUInt32(bytes, changed.Offset, replacement);
        }, Newsletter);

        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input));
    }

    private static PublisherTextFrame[] Frames(PublisherDocument document, uint storyId) => document.Pages.Concat(document.MasterPages)
        .SelectMany(page => page.TextFrames).Where(frame => frame.StoryId == storyId).OrderBy(frame => frame.Order).ToArray();
}

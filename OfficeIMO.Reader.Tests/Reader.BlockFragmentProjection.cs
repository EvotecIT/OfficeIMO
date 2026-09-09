using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderBlockFragmentProjectionTests {
    [Theory]
    [InlineData("inherit")]
    [InlineData("conflict")]
    [InlineData("aggregate")]
    public void GeometryFallbackRequiresOneCompatibleObservation(string mode) {
        var aggregate = new OfficeDocumentBlock { Id = "source", Text = "Aggregate text", Region = mode == "aggregate" ? new() { X = 9 } : null };
        var source = new OfficeDocumentReadResult { Blocks = new[] { aggregate }, Pages = new[] { new OfficeDocumentPage { Number = 1,
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "source", Text = "Page text", Region = new() { X = 1 } },
                new OfficeDocumentBlock { Id = "source", Text = "Page text", Region = new() { X = mode == "conflict" ? 2 : 1 } }
            } } } };
        foreach (var document in new[] { source, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) }) {
            var block = Assert.Single(document.EnumerateBlocks());
            Assert.Equal("Aggregate text", block.Text);
            if (mode == "conflict") Assert.Null(block.Region);
            else Assert.Equal(mode == "aggregate" ? 9 : 1, block.Region!.X);
        }
        Assert.Equal(mode == "aggregate" ? 9 : (double?)null, aggregate.Region?.X);
    }

    [Theory]
    [InlineData(5)]
    [InlineData(1200)]
    public void WordParagraphRetainsComputedPageTextAndGeometry(int words) {
        using var stream = new MemoryStream();
        using (var word = OfficeIMO.Word.WordDocument.Create(stream)) {
            word.AddParagraph(string.Join(" ", Enumerable.Range(0, words).Select(index => "word" + index)));
            word.Save();
        }
        stream.Position = 0;
        var source = OfficeIMO.Reader.Tests.ReaderTestReaders.Word(includePageLocations: true).ReadDocument(stream, "paragraph.docx");
        var fragments = source.Pages.SelectMany(page => page.Blocks).Where(block => !string.IsNullOrWhiteSpace(block.Text)).ToArray();
        if (words > 5) Assert.True(fragments.Length > 1);
        else Assert.Single(fragments);
        Assert.Single(fragments.Select(block => block.Id).Distinct());
        foreach (var document in new[] { source, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) }) {
            var blocks = document.EnumerateBlocks().Where(block => !string.IsNullOrWhiteSpace(block.Text)).ToArray();
            Assert.Equal(fragments.Select(block => block.Text), blocks.Select(block => block.Text));
            Assert.Equal(fragments.Select(block => block.Location!.Page), blocks.Select(block => block.Location!.Page));
            Assert.All(blocks, block => Assert.NotNull(block.Region));
        }
    }

    [Fact]
    public void PartialPageInspectionCannotLocateTheFullAggregateAtItsFirstFragment() {
        var source = new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Id = "source", Text = "Alpha. Beta." } },
            Pages = new[] {
                new OfficeDocumentPage { Number = 1, Blocks = new[] { new OfficeDocumentBlock { Id = "source", Text = "Alpha." } }
                    .Concat(Enumerable.Range(0, 14).Select(index => new OfficeDocumentBlock { Id = "filler" + index, Text = "Other" })).ToArray() },
                new OfficeDocumentPage { Number = 2, Blocks = new[] { new OfficeDocumentBlock { Id = "source", Text = "Beta." } } }
            } };
        var hierarchy = ReaderHierarchicalChunker.Chunk(source, new ReaderHierarchicalChunkingOptions {
            MaxInputChunks = 1, MaxTokens = 100, OverlapTokens = 0, IncludeContextInText = false });
        var chunk = Assert.Single(hierarchy.Chunks);
        Assert.Equal("Alpha.", chunk.Text);
        Assert.Equal(1, chunk.Location.Page);
        Assert.Contains(hierarchy.Diagnostics, item => item.Code == "hierarchical-input-chunk-limit");
    }

    [Theory]
    [InlineData("page", true)]
    [InlineData("slide", true)]
    [InlineData("sheet", true)]
    [InlineData("path", true)]
    public void SharedSourceIdentityRetainsContainerFragments(string container, bool useId) {
        ReaderLocation Location(int? number) => new() { BlockAnchor = "paragraph", Page = container == "page" ? number : null,
            Slide = container == "slide" ? number : null, Sheet = container == "sheet" && number.HasValue ? "Sheet" + number : null,
            Path = container == "path" && number.HasValue ? "file" + number : null };
        OfficeDocumentBlock Block(string text, int? number) => new() { Id = useId ? "paragraph" : "", Kind = "paragraph", Text = text,
            Location = Location(number), Region = number.HasValue ? new() { X = number.Value * 10, Y = 1, Width = 2, Height = 3 } : null };
        var aggregate = Block("First fragment. Second fragment.", null);
        var source = new OfficeDocumentReadResult { Blocks = new[] { aggregate }, Pages = new[] {
            new OfficeDocumentPage { Location = Location(1), Blocks = new[] { Block("First fragment.", 1) } },
            new OfficeDocumentPage { Location = Location(2), Blocks = new[] { Block("Second fragment.", 2) } }
        } };
        foreach (var document in new[] { source, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) }) {
            var blocks = document.EnumerateBlocks().ToArray();
            Assert.Equal(new[] { "First fragment.", "Second fragment." }, blocks.Select(block => block.Text));
            Assert.Equal(new double[] { 10, 20 }, blocks.Select(block => block.Region!.X));
            Assert.Equal(blocks.Select(block => block.Text), document.EnumerateContent().Select(item => item.Block!.Text));
            var hierarchy = ReaderHierarchicalChunker.Chunk(document, new ReaderHierarchicalChunkingOptions {
                MaxInputChunks = 10, MaxTokens = 100, OverlapTokens = 0, IncludeContextInText = false });
            Assert.Equal(new[] { "First fragment.", "Second fragment." }, hierarchy.Chunks.Select(chunk => chunk.Text));
        }
        Assert.Null(aggregate.Location!.Page);
    }
}

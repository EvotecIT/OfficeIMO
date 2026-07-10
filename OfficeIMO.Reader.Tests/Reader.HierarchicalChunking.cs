using OfficeIMO.Reader;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderHierarchicalChunkingTests {
    private static readonly IReaderTokenCounter WordCounter = new WordTokenCounter();

    [Fact]
    public void Chunk_SplitsToHardTokenBoundsWithDeterministicOverlapAndIds() {
        ReaderChunk source = CreateChunk(
            "source-1",
            "one two three four five six seven eight nine ten",
            page: 1,
            headingPath: "Root > Child");
        var options = new ReaderHierarchicalChunkingOptions {
            MaxTokens = 4,
            OverlapTokens = 1,
            IncludeContextInText = false,
            TokenCounter = WordCounter
        };

        ReaderChunkHierarchyResult first = ReaderHierarchicalChunker.Chunk(new[] { source }, options);
        ReaderChunkHierarchyResult second = ReaderHierarchicalChunker.Chunk(new[] { source }, options);

        Assert.True(first.Chunks.Count > 1);
        Assert.All(first.Chunks, chunk => Assert.InRange(chunk.TokenEstimate ?? -1, 1, 4));
        Assert.Equal(first.Chunks.Select(chunk => chunk.Id), second.Chunks.Select(chunk => chunk.Id));
        Assert.Equal(first.Chunks.Select(chunk => chunk.Text), second.Chunks.Select(chunk => chunk.Text));
        Assert.Equal("one two three four five six seven eight nine ten", source.Text);
        Assert.Null(source.ChunkHash);

        for (int index = 1; index < first.Segments.Count; index++) {
            Assert.True(first.Segments[index].StartCharacter < first.Segments[index - 1].EndCharacter);
            Assert.InRange(first.Segments[index].OverlapTokenCount, 1, 1);
        }
        Assert.Equal(0, first.Segments[0].StartCharacter);
        Assert.Equal(source.Text.Length, first.Segments[first.Segments.Count - 1].EndCharacter);
        Assert.All(first.Segments, segment => Assert.Equal(
            source.Text.Substring(segment.StartCharacter, segment.EndCharacter - segment.StartCharacter),
            first.Chunks[segment.SegmentIndex].Text));
        Assert.Equal(first.OutputTokenCount, first.Chunks.Sum(chunk => chunk.TokenEstimate ?? 0));
        Assert.Equal(first.OverlapTokenCount, first.Segments.Sum(segment => segment.OverlapTokenCount));
    }

    [Fact]
    public void Chunk_BuildsDocumentContainerHeadingAndLeafHierarchy() {
        var document = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { SourceId = "doc-1", Path = "guide.pdf", Title = "Guide" },
            Chunks = new[] {
                CreateChunk("c1", "alpha beta", page: 1, headingPath: "Root > Child", headingSlug: "child"),
                CreateChunk("c2", "gamma delta", page: 1, headingPath: "Root > Child", headingSlug: "child"),
                CreateChunk("c3", "epsilon", page: 2, headingPath: "Appendix", headingSlug: "appendix")
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
            document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 100,
                OverlapTokens = 0,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        ReaderChunkHierarchyNode root = Assert.Single(result.Nodes, node => node.Kind == ReaderChunkHierarchyNodeKind.Document);
        Assert.Equal("Guide", root.Title);
        Assert.Equal(result.RootNodeId, root.Id);
        Assert.Equal(2, root.ChildNodeIds.Count);
        Assert.Equal(result.OutputTokenCount, root.TokenCount);
        Assert.Equal(2, result.Nodes.Count(node => node.Kind == ReaderChunkHierarchyNodeKind.Container));
        Assert.Equal(3, result.Nodes.Count(node => node.Kind == ReaderChunkHierarchyNodeKind.Heading));
        Assert.Equal(3, result.Nodes.Count(node => node.Kind == ReaderChunkHierarchyNodeKind.Chunk));

        ReaderChunkHierarchyNode child = Assert.Single(result.Nodes, node => node.Kind == ReaderChunkHierarchyNodeKind.Heading && node.Title == "Child");
        Assert.Equal(2, child.ChildNodeIds.Count);
        Assert.Equal(4L, child.TokenCount);
    }

    [Fact]
    public void Chunk_PreservesStructuredSidecarsOnceAcrossSplitSegments() {
        var table = new ReaderTable { Title = "Data", Columns = new[] { "A" } };
        ReaderChunk source = CreateChunk("source", "one two three four five six");
        source.Tables = new[] { table };
        source.Warnings = new[] { "source warning" };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
            new[] { source },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 2,
                OverlapTokens = 0,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.True(result.Chunks.Count > 1);
        Assert.Same(table, Assert.Single(Assert.IsAssignableFrom<IReadOnlyList<ReaderTable>>(result.Chunks[0].Tables)));
        Assert.Equal("source warning", Assert.Single(result.Chunks[0].Warnings!));
        Assert.All(result.Chunks.Skip(1), chunk => {
            Assert.Null(chunk.Tables);
            Assert.Null(chunk.Warnings);
        });
    }

    [Fact]
    public void Chunk_AddsContextWithinBudgetOrRetainsItAsMetadata() {
        ReaderChunk source = CreateChunk("source", "alpha beta gamma", page: 3, headingPath: "Root > Child");
        var withContext = new ReaderHierarchicalChunkingOptions {
            MaxTokens = 10,
            OverlapTokens = 0,
            IncludeContextInText = true,
            TokenCounter = WordCounter
        };

        ReaderChunkHierarchyResult included = ReaderHierarchicalChunker.Chunk(new[] { source }, withContext);

        Assert.StartsWith("Page 3 > Root > Child\n\n", Assert.Single(included.Chunks).Text, StringComparison.Ordinal);
        Assert.True(Assert.Single(included.Segments).ContextTokenCount > 0);
        Assert.InRange(Assert.Single(included.Chunks).TokenEstimate ?? -1, 1, 10);

        withContext.MaxTokens = 3;
        ReaderChunkHierarchyResult omitted = ReaderHierarchicalChunker.Chunk(new[] { source }, withContext);
        Assert.DoesNotContain("Page 3", omitted.Chunks[0].Text, StringComparison.Ordinal);
        Assert.Equal("Page 3 > Root > Child", omitted.Segments[0].Context);
        Assert.Contains(omitted.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-context-omitted");
    }

    [Fact]
    public void Chunk_EnforcesInputOutputAndDepthBoundsWithDiagnostics() {
        ReaderChunk[] chunks = {
            CreateChunk("c1", "one two three four five six", headingPath: "A > B > C > D"),
            CreateChunk("c2", "seven eight")
        };
        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
            chunks,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 2,
                OverlapTokens = 0,
                MaxInputChunks = 1,
                MaxOutputChunks = 2,
                MaxHierarchyDepth = 2,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.Equal(2, result.Chunks.Count);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-output-chunk-limit");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-depth-limit");
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");

        ReaderChunkHierarchyResult inputBound = ReaderHierarchicalChunker.Chunk(
            chunks,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 100,
                OverlapTokens = 0,
                MaxInputChunks = 1,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });
        Assert.Single(inputBound.Chunks);
        Assert.Contains(inputBound.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");
    }

    [Fact]
    public void Chunk_IsDeterministicJsonAndRejectsInvalidContracts() {
        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
            new[] { CreateChunk("source", "alpha beta") },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.Equal(result.ToJson(), ReaderHierarchicalChunker.Chunk(
            new[] { CreateChunk("source", "alpha beta") },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            }).ToJson());
        using JsonDocument json = JsonDocument.Parse(result.ToJson());
        Assert.Equal(ReaderChunkHierarchySchema.Id, json.RootElement.GetProperty("schemaId").GetString());
        Assert.Equal("Document", json.RootElement.GetProperty("nodes")[0].GetProperty("kind").GetString());

        Assert.Throws<ArgumentOutOfRangeException>(() => ReaderHierarchicalChunker.Chunk(
            Array.Empty<ReaderChunk>(),
            new ReaderHierarchicalChunkingOptions { MaxTokens = 1, OverlapTokens = 1 }));
        Assert.Throws<InvalidOperationException>(() => new ReaderChunkHierarchyResult { SchemaVersion = 2 }.ToJson());
    }

    [Fact]
    public void Chunk_HonorsCancellationAndValidatesTokenCounters() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => ReaderHierarchicalChunker.Chunk(
            new[] { CreateChunk("source", "alpha") },
            cancellationToken: cancellation.Token));
        Assert.Throws<InvalidOperationException>(() => ReaderHierarchicalChunker.Chunk(
            new[] { CreateChunk("source", "alpha") },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                TokenCounter = new NegativeTokenCounter()
            }));
    }

    [Fact]
    public void Chunk_FallsBackToRichBlocksAndInfersHeadingPaths() {
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Word,
            Source = new OfficeDocumentSource { SourceId = "rich-document", Title = "Rich" },
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "root", Kind = "heading", Level = 1, Text = "Root" },
                new OfficeDocumentBlock { Id = "body", Kind = "paragraph", Text = "Root body" },
                new OfficeDocumentBlock { Id = "child", Kind = "heading", Level = 2, Text = "Child" },
                new OfficeDocumentBlock { Id = "child-body", Kind = "paragraph", Text = "Child body" }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
            document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 100,
                OverlapTokens = 0,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.Equal(4, result.Chunks.Count);
        Assert.Contains(result.Nodes, node => node.Kind == ReaderChunkHierarchyNodeKind.Heading && node.Title == "Root");
        Assert.Contains(result.Nodes, node => node.Kind == ReaderChunkHierarchyNodeKind.Heading && node.Title == "Child");
        Assert.Equal("Root > Child", result.Segments[3].Context);
        Assert.Equal("paragraph", result.Chunks[3].Location.SourceBlockKind);
    }

    [Fact]
    public void Chunk_BoundsContextAndRejectsMixedSourceCollections() {
        ReaderChunk source = CreateChunk("source", "alpha beta", headingPath: "abcdefghij");
        ReaderChunkHierarchyResult bounded = ReaderHierarchicalChunker.Chunk(
            new[] { source },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxContextCharacters = 5,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.Equal("abcde", Assert.Single(bounded.Segments).Context);
        Assert.Contains(bounded.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-context-character-limit");

        source.Location.HeadingPath = "😀abcdef";
        bounded = ReaderHierarchicalChunker.Chunk(
            new[] { source },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxContextCharacters = 2,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });
        Assert.Equal("😀", Assert.Single(bounded.Segments).Context);

        ReaderChunk other = CreateChunk("other", "gamma");
        other.SourceId = "another-source";
        Assert.Throws<InvalidOperationException>(() => ReaderHierarchicalChunker.Chunk(
            new[] { source, other },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                TokenCounter = WordCounter
            }));
    }

    [Fact]
    public void Chunk_MakesProgressAcrossUnicodeAndSmallHeuristicBudgets() {
        const string text = "alpha 😀 beta. gamma\n\ndelta 😀 epsilon zeta eta theta";
        for (int maxTokens = 1; maxTokens <= 16; maxTokens++) {
            ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
                new[] { CreateChunk("unicode", text) },
                new ReaderHierarchicalChunkingOptions {
                    MaxTokens = maxTokens,
                    OverlapTokens = Math.Min(2, maxTokens - 1),
                    IncludeContextInText = false
                });

            Assert.NotEmpty(result.Chunks);
            Assert.All(result.Chunks, chunk => Assert.InRange(chunk.TokenEstimate ?? -1, 0, maxTokens));
            Assert.Equal(0, result.Segments[0].StartCharacter);
            Assert.Equal(text.Length, result.Segments[result.Segments.Count - 1].EndCharacter);
            for (int index = 1; index < result.Segments.Count; index++) {
                Assert.True(result.Segments[index].StartCharacter <= result.Segments[index - 1].EndCharacter);
                Assert.True(result.Segments[index].StartCharacter > result.Segments[index - 1].StartCharacter);
            }
        }
    }

    [Fact]
    public async Task StaticAndInstanceReadSurfacesSupportSyncAsyncAndProcessors() {
        byte[] markdown = Encoding.UTF8.GetBytes("# Overview\n\nUseful body");
        var options = new ReaderHierarchicalChunkingOptions {
            MaxTokens = 100,
            OverlapTokens = 0,
            IncludeContextInText = false
        };

        ReaderChunkHierarchyResult staticSync = DocumentReader.ReadHierarchical(markdown, "note.md", chunkingOptions: options);
        ReaderChunkHierarchyResult staticAsync = await DocumentReader.ReadHierarchicalAsync(markdown, "note.md", chunkingOptions: options);
        Assert.Equal(staticSync.ToJson(), staticAsync.ToJson());

        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddProcessor(new DelegateOfficeDocumentProcessor("prefix", (document, _) => {
                foreach (ReaderChunk chunk in document.Chunks) {
                    chunk.Text = "processed " + chunk.Text;
                    chunk.Markdown = null;
                }
                return document;
            }))
            .Build();
        ReaderChunkHierarchyResult instanceSync = reader.ReadHierarchical(markdown, "note.md", chunkingOptions: options);
        ReaderChunkHierarchyResult instanceAsync = await reader.ReadHierarchicalAsync(markdown, "note.md", chunkingOptions: options);
        Assert.StartsWith("processed ", Assert.Single(instanceSync.Chunks).Text, StringComparison.Ordinal);
        Assert.Equal(instanceSync.ToJson(), instanceAsync.ToJson());
    }

    private static ReaderChunk CreateChunk(
        string id,
        string text,
        int? page = null,
        string? headingPath = null,
        string? headingSlug = null) => new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Text,
            SourceId = "source-document",
            Text = text,
            Location = new ReaderLocation {
                Path = "source.txt",
                Page = page,
                HeadingPath = headingPath,
                HeadingSlug = headingSlug,
                BlockAnchor = id
            }
        };

    private sealed class WordTokenCounter : IReaderTokenCounter {
        public string Id => "tests.words-v1";

        public int CountTokens(string text) => string.IsNullOrWhiteSpace(text)
            ? 0
            : text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
    }

    private sealed class NegativeTokenCounter : IReaderTokenCounter {
        public string Id => "tests.negative-v1";
        public int CountTokens(string text) => -1;
    }
}

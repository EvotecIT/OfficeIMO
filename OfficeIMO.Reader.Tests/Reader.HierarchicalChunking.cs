using OfficeIMO.Reader;
using OfficeIMO.Reader.Markdown;
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
    public void Chunk_FallsBackToTextWhenPreferredMarkdownIsEmpty() {
        ReaderChunk source = CreateChunk("source", "plain text body");
        source.Markdown = string.Empty;

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
            new[] { source },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                IncludeContextInText = false,
                PreferMarkdown = true,
                TokenCounter = WordCounter
            });

        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.Equal("plain text body", chunk.Text);
        Assert.Null(chunk.Markdown);
    }

    [Fact]
    public void Chunk_PreservesFullHeadingIdentityWhenDisplayTitlesAreTruncated() {
        ReaderChunk first = CreateChunk("first", "one", headingPath: "ABCDE-first");
        ReaderChunk second = CreateChunk("second", "two", headingPath: "ABCDE-second");

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
            new[] { first, second },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxContextCharacters = 5,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        ReaderChunkHierarchyNode[] headings = result.Nodes
            .Where(node => node.Kind == ReaderChunkHierarchyNodeKind.Heading)
            .ToArray();
        Assert.Equal(2, headings.Length);
        Assert.All(headings, heading => Assert.Equal("ABCDE", heading.Title));
        Assert.Equal(2, headings.Select(heading => heading.Id).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void MarkdownRead_PreservesRawHeadingPathWhileHierarchyKeepsLiteralDelimiter() {
        byte[] markdown = Encoding.UTF8.GetBytes("# Q1 > Q2\\Back\n\nBody");
        ReaderChunk normal = Assert.Single(OfficeIMO.Reader.Tests.ReaderTestReaders.All.Read(markdown, "note.md"));

        Assert.Equal("Q1 > Q2\\Back", normal.Location.HeadingPath);

        ReaderChunkHierarchyResult hierarchy = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadHierarchical(
            markdown,
            "note.md",
            chunkingOptions: new ReaderHierarchicalChunkingOptions {
                MaxTokens = 100,
                OverlapTokens = 0,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });
        ReaderChunkHierarchyNode heading = Assert.Single(
            hierarchy.Nodes,
            node => node.Kind == ReaderChunkHierarchyNodeKind.Heading);
        Assert.Equal("Q1 > Q2\\Back", heading.Title);
    }

    [Fact]
    public void Chunk_UsesProcessorUpdatedPublicHeadingPathInsteadOfStaleHint() {
        ReaderChunk chunk = Assert.Single(OfficeIMO.Reader.Tests.ReaderTestReaders.All.Read(
            Encoding.UTF8.GetBytes("# Original > Literal\n\nBody"),
            "note.md"));
        chunk.Location.HeadingPath = "Processed";

        ReaderChunkHierarchyResult hierarchy = ReaderHierarchicalChunker.Chunk(
            new[] { chunk },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 100,
                OverlapTokens = 0,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        ReaderChunkHierarchyNode heading = Assert.Single(
            hierarchy.Nodes,
            node => node.Kind == ReaderChunkHierarchyNodeKind.Heading);
        Assert.Equal("Processed", heading.Title);
    }

    [Fact]
    public void Chunk_InheritsMissingEnvelopeProvenanceAndReplacesWhitespacePath() {
        ReaderChunk source = CreateChunk("source", "body");
        source.SourceId = "chunk-id";
        source.SourceHash = null;
        source.SourceLastWriteUtc = null;
        source.SourceLengthBytes = null;
        source.Location.Path = "   ";
        DateTime lastWrite = new DateTime(2026, 7, 11, 1, 2, 3, DateTimeKind.Utc);
        var document = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource {
                SourceId = "envelope-id",
                Path = "source.md",
                SourceHash = "source-hash",
                LastWriteUtc = lastWrite,
                LengthBytes = 123
            },
            Chunks = new[] { source }
        };

        ReaderChunk inherited = Assert.Single(ReaderHierarchicalChunker.Chunk(
            document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            }).Chunks);

        Assert.Equal("chunk-id", inherited.SourceId);
        Assert.Equal("source.md", inherited.Location.Path);
        Assert.Equal("source-hash", inherited.SourceHash);
        Assert.Equal(lastWrite, inherited.SourceLastWriteUtc);
        Assert.Equal(123, inherited.SourceLengthBytes);
    }

    [Fact]
    public void Chunk_SeparatesDelimitedHeadingKeysAndRootIdentityNamespaces() {
        ReaderChunk literalMarker = CreateChunk("literal", "one", headingPath: "A|slug:x");
        ReaderChunk slugIdentity = CreateChunk("slugged", "two", headingPath: "A", headingSlug: "x");
        var options = new ReaderHierarchicalChunkingOptions {
            MaxTokens = 10,
            OverlapTokens = 0,
            IncludeContextInText = false,
            TokenCounter = WordCounter
        };

        ReaderChunkHierarchyResult headings = ReaderHierarchicalChunker.Chunk(
            new[] { literalMarker, slugIdentity },
            options);
        Assert.Equal(2, headings.Nodes.Count(node => node.Kind == ReaderChunkHierarchyNodeKind.Heading));

        var sourceIdDocument = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { SourceId = "foo" },
            Chunks = new[] { CreateChunk("source", "body") }
        };
        sourceIdDocument.Chunks[0].SourceId = "foo";
        sourceIdDocument.Chunks[0].Location.Path = null;
        var pathDocument = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "foo" },
            Chunks = new[] { CreateChunk("source", "body") }
        };
        pathDocument.Chunks[0].SourceId = null;
        pathDocument.Chunks[0].Location.Path = "foo";

        Assert.NotEqual(
            ReaderHierarchicalChunker.Chunk(sourceIdDocument, options).RootNodeId,
            ReaderHierarchicalChunker.Chunk(pathDocument, options).RootNodeId);
    }

    [Fact]
    public void Chunk_ReportsInputLimitOnlyWhenAnotherChunkExists() {
        var options = new ReaderHierarchicalChunkingOptions {
            MaxTokens = 10,
            OverlapTokens = 0,
            MaxInputChunks = 1,
            IncludeContextInText = false,
            TokenCounter = WordCounter
        };

        ReaderChunkHierarchyResult exact = ReaderHierarchicalChunker.Chunk(
            new[] { CreateChunk("first", "first block") },
            options);
        ReaderChunkHierarchyResult truncated = ReaderHierarchicalChunker.Chunk(
            new[] { CreateChunk("first", "first block"), CreateChunk("second", "second block") },
            options);

        Assert.DoesNotContain(exact.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");
        Assert.Contains(truncated.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");
    }

    [Fact]
    public void Chunk_BoundsFallbackInspectionWhenSourceBlocksAreDuplicates() {
        var duplicate = new OfficeDocumentBlock { Id = "duplicate", Text = "body" };
        var blocks = new CountingBlockList(duplicate, 1000);
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Blocks = blocks
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxInputChunks = 2,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.Single(result.Chunks);
        Assert.Equal(8, blocks.ReadCount);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");
    }

    [Fact]
    public void Chunk_DoesNotChargeDuplicateFallbackBlocksToInputQuota() {
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "first", Text = "first" },
                new OfficeDocumentBlock { Id = "first", Text = "duplicate" },
                new OfficeDocumentBlock { Id = "second", Text = "second" }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxInputChunks = 2,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.Equal(new[] { "first", "second" }, result.Chunks.Select(chunk => chunk.Text));
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");
    }

    [Fact]
    public void Chunk_ReportsInputLimitWhenFallbackBlocksAreTruncated() {
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "first", Text = "first" },
                new OfficeDocumentBlock { Id = "second", Text = "second" }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxInputChunks = 1,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.Single(result.Chunks);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");
    }

    [Fact]
    public void Chunk_Does_Not_Report_Input_Truncation_For_Duplicate_Page_Blocks() {
        var block = new OfficeDocumentBlock { Id = "same", Text = "body" };
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Blocks = new[] { block },
            Pages = new[] { new OfficeDocumentPage { Number = 1, Blocks = new[] { block } } }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxInputChunks = 1,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.Single(result.Chunks);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");
    }

    [Fact]
    public void Chunk_InheritsPageLocationBeyondInputChunkOrdinal() {
        var block = new OfficeDocumentBlock { Id = "retained", Text = "body" };
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Blocks = new[] { block },
            Pages = new[] {
                new OfficeDocumentPage { Number = 1, Blocks = Array.Empty<OfficeDocumentBlock>() },
                new OfficeDocumentPage { Number = 2, Blocks = new[] { block } }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxInputChunks = 1,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.Equal(2, chunk.Location.Page);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");
    }

    [Fact]
    public void Chunk_InheritsPageLocationPastUnrelatedPageBlocks() {
        var retained = new OfficeDocumentBlock { Id = "retained", Text = "body" };
        var unrelated = new OfficeDocumentBlock { Id = "unrelated", Text = "other" };
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Blocks = new[] { retained },
            Pages = new[] {
                new OfficeDocumentPage { Number = 1, Blocks = new[] { unrelated } },
                new OfficeDocumentPage { Number = 2, Blocks = new[] { retained } }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxInputChunks = 1,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.Equal(2, chunk.Location.Page);
    }

    [Fact]
    public void Chunk_InheritsPageLocationBeyondUnrelatedInspectionAllowance() {
        var retained = new OfficeDocumentBlock { Id = "retained", Text = "body" };
        var unrelated = Enumerable.Range(0, 5)
            .Select(index => new OfficeDocumentBlock { Id = "unrelated-" + index, Text = "other" })
            .ToArray();
        var pageBlocks = unrelated.Concat(new[] { retained }).ToArray();
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Blocks = new[] { retained },
            Pages = new[] { new OfficeDocumentPage { Number = 7, Blocks = pageBlocks } }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxInputChunks = 1,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.Equal(7, chunk.Location.Page);
    }

    [Fact]
    public void Chunk_BoundsPageIndexInspectionWhenRetainedBlocksAreMissing() {
        var retained = new OfficeDocumentBlock { Id = "retained", Text = "body" };
        var unrelated = new OfficeDocumentBlock { Id = "unrelated", Text = "other" };
        var pageBlocks = new CountingBlockList(unrelated, 1_000);
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Blocks = new[] {
                retained,
                new OfficeDocumentBlock { Id = "truncated", Text = "later" }
            },
            Pages = new[] { new OfficeDocumentPage { Number = 7, Blocks = pageBlocks } }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxInputChunks = 1,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.Single(result.Chunks);
        Assert.Equal(15, pageBlocks.ReadCount);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");
    }

    [Fact]
    public void Chunk_BoundsPageIndexInspectionAcrossEmptyPages() {
        var retained = new OfficeDocumentBlock { Id = "retained", Text = "body" };
        var pages = new CountingPageList(
            new OfficeDocumentPage { Number = 1, Blocks = Array.Empty<OfficeDocumentBlock>() },
            1_000);
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Blocks = new[] {
                retained,
                new OfficeDocumentBlock { Id = "truncated", Text = "later" }
            },
            Pages = pages
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxInputChunks = 1,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.Single(result.Chunks);
        Assert.Equal(16, pages.ReadCount);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");
    }

    [Fact]
    public void Chunk_BoundsFallbackInspectionAcrossEmptyPages() {
        var retained = new OfficeDocumentBlock { Id = "retained", Text = "body" };
        var pages = new CountingPageList(
            new OfficeDocumentPage { Number = 1, Blocks = Array.Empty<OfficeDocumentBlock>() },
            1_000);
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Blocks = new[] { retained },
            Pages = pages
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                MaxInputChunks = 1,
                IncludeContextInText = false,
                TokenCounter = WordCounter
            });

        Assert.Single(result.Chunks);
        Assert.Equal(20, pages.ReadCount);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-input-chunk-limit");
    }

    [Fact]
    public void Chunk_BoundsLeafTitlesAndPathsWithoutChangingLeafIdentity() {
        ReaderChunk source = CreateChunk("stable-id", "body");
        source.Location.Page = null;
        source.Location.BlockAnchor = new string('a', 50);
        var options = new ReaderHierarchicalChunkingOptions {
            MaxTokens = 10,
            OverlapTokens = 0,
            MaxContextCharacters = 5,
            IncludeContextInText = false,
            TokenCounter = WordCounter
        };

        ReaderChunkHierarchyResult bounded = ReaderHierarchicalChunker.Chunk(new[] { source }, options);
        ReaderChunkHierarchyNode boundedLeaf = Assert.Single(
            bounded.Nodes,
            node => node.Kind == ReaderChunkHierarchyNodeKind.Chunk);
        options.MaxContextCharacters = 10;
        ReaderChunkHierarchyNode widerLeaf = Assert.Single(
            ReaderHierarchicalChunker.Chunk(new[] { source }, options).Nodes,
            node => node.Kind == ReaderChunkHierarchyNodeKind.Chunk);

        Assert.Equal(5, boundedLeaf.Title.Length);
        Assert.InRange(boundedLeaf.Path.Length, 1, 5);
        Assert.Equal(boundedLeaf.Id, widerLeaf.Id);
    }

    [Fact]
    public void Chunk_OmitsContextWhenItLeavesNoRoomForFirstSourceToken() {
        ReaderChunk source = CreateChunk("source", "a", headingPath: "Context");

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
            new[] { source },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 2,
                OverlapTokens = 0,
                IncludeContextInText = true,
                TokenCounter = new NonAdditiveContextTokenCounter()
            });

        Assert.Equal("a", Assert.Single(result.Chunks).Text);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "hierarchical-context-omitted");
    }

    [Fact]
    public void Chunk_UsesPathIdentityForSegmentIdsWhenSourceIdIsMissing() {
        ReaderChunk first = CreateChunk("same", "body");
        first.SourceId = null;
        first.Location.Path = "first.txt";
        ReaderChunk second = CreateChunk("same", "body");
        second.SourceId = null;
        second.Location.Path = "second.txt";
        var options = new ReaderHierarchicalChunkingOptions {
            MaxTokens = 10,
            OverlapTokens = 0,
            IncludeContextInText = false,
            TokenCounter = WordCounter
        };

        string firstId = Assert.Single(ReaderHierarchicalChunker.Chunk(new[] { first }, options).Chunks).Id;
        string secondId = Assert.Single(ReaderHierarchicalChunker.Chunk(new[] { second }, options).Chunks).Id;

        Assert.NotEqual(firstId, secondId);
    }

    [Fact]
    public void Chunk_InheritsEnvelopePathForExistingChunksWithoutSourceIdentity() {
        ReaderChunk source = CreateChunk("same", "body");
        source.SourceId = null;
        source.Location.Path = null;
        var first = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "first.txt" },
            Chunks = new[] { source }
        };
        var second = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "second.txt" },
            Chunks = new[] { source }
        };
        var options = new ReaderHierarchicalChunkingOptions {
            MaxTokens = 10,
            OverlapTokens = 0,
            IncludeContextInText = false,
            TokenCounter = WordCounter
        };

        ReaderChunk firstChunk = Assert.Single(ReaderHierarchicalChunker.Chunk(first, options).Chunks);
        ReaderChunk secondChunk = Assert.Single(ReaderHierarchicalChunker.Chunk(second, options).Chunks);

        Assert.Equal("first.txt", firstChunk.Location.Path);
        Assert.NotEqual(firstChunk.Id, secondChunk.Id);
        Assert.Null(source.Location.Path);
    }

    [Fact]
    public void Chunk_UsesEnvelopePathForFallbackBlockIdentity() {
        var block = new OfficeDocumentBlock { Id = "same", Kind = "paragraph", Text = "body" };
        var first = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "first.txt" },
            Blocks = new[] { block }
        };
        var second = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "second.txt" },
            Blocks = new[] { block }
        };

        ReaderChunk firstChunk = Assert.Single(ReaderHierarchicalChunker.Chunk(first).Chunks);
        ReaderChunk secondChunk = Assert.Single(ReaderHierarchicalChunker.Chunk(second).Chunks);

        Assert.Equal("first.txt", firstChunk.Location.Path);
        Assert.NotEqual(firstChunk.Id, secondChunk.Id);
    }

    [Fact]
    public void Chunk_InheritsPageContainerForPageOnlyBlocks() {
        var document = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "pages.pdf" },
            Pages = new[] {
                new OfficeDocumentPage {
                    Number = 4,
                    Blocks = new[] { new OfficeDocumentBlock { Id = "p4", Kind = "paragraph", Text = "Page body" } }
                }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document);

        Assert.Equal(4, Assert.Single(result.Chunks).Location.Page);
        ReaderChunkHierarchyNode page = Assert.Single(result.Nodes, node => node.Kind == ReaderChunkHierarchyNodeKind.Container);
        Assert.Equal("Page 4", page.Title);
    }

    [Fact]
    public void Chunk_PrefersPageScopeForBlocksAlsoPresentAtDocumentLevel() {
        var documentBlock = new OfficeDocumentBlock { Id = "shared", Kind = "paragraph", Text = "Page body" };
        var pageClone = new OfficeDocumentBlock { Id = "shared", Kind = "paragraph", Text = "Page body" };
        var document = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "pages.pdf" },
            Blocks = new[] { documentBlock },
            Pages = new[] {
                new OfficeDocumentPage { Number = 7, Blocks = new[] { pageClone } }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document);

        Assert.Equal(7, Assert.Single(result.Chunks).Location.Page);
        Assert.Equal("Page 7", Assert.Single(result.Nodes, node => node.Kind == ReaderChunkHierarchyNodeKind.Container).Title);
    }

    [Fact]
    public void Chunk_UsesSheetContainerKindForBlockOnlyResults() {
        var document = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "workbook.xlsx" },
            Pages = new[] {
                new OfficeDocumentPage {
                    Number = 2,
                    Name = "Inventory",
                    Location = new ReaderLocation { SourceBlockKind = "sheet" },
                    Blocks = new[] { new OfficeDocumentBlock { Id = "sheet-body", Kind = "paragraph", Text = "Sheet body" } }
                }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document);

        ReaderLocation location = Assert.Single(result.Chunks).Location;
        Assert.Equal("Inventory", location.Sheet);
        Assert.Null(location.Page);
        Assert.Equal("Sheet: Inventory", Assert.Single(result.Nodes, node => node.Kind == ReaderChunkHierarchyNodeKind.Container).Title);
    }

    [Fact]
    public void Chunk_InheritsSheetContainerWhenBlockSheetIsWhitespace() {
        var document = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "workbook.xlsx" },
            Pages = new[] {
                new OfficeDocumentPage {
                    Number = 2,
                    Name = "Inventory",
                    Location = new ReaderLocation { SourceBlockKind = "sheet" },
                    Blocks = new[] {
                        new OfficeDocumentBlock {
                            Id = "sheet-body",
                            Kind = "paragraph",
                            Text = "Sheet body",
                            Location = new ReaderLocation { Sheet = " " }
                        }
                    }
                }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document);

        ReaderLocation location = Assert.Single(result.Chunks).Location;
        Assert.Equal("Inventory", location.Sheet);
        Assert.Null(location.Page);
    }

    [Fact]
    public void Chunk_PreservesLiteralHeadingSeparators() {
        var document = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { SourceId = "literal-heading" },
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "quarter", Kind = "heading", Level = 1, Text = "Q1 > Q2" },
                new OfficeDocumentBlock { Id = "body", Kind = "paragraph", Text = "Comparison" }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document);

        ReaderChunkHierarchyNode heading = Assert.Single(result.Nodes, node => node.Kind == ReaderChunkHierarchyNodeKind.Heading);
        Assert.Equal("Q1 > Q2", heading.Title);
        Assert.Equal("Q1 > Q2", result.Segments[1].Context);
        Assert.Equal("Q1 > Q2", result.Chunks[1].Location.HeadingPath);
    }

    [Fact]
    public void Chunk_PreservesRawEscapesInContextWithoutHierarchyMetadata() {
        ReaderChunk source = CreateChunk("raw-context", "Body", headingPath: @"Q1 \> Q2\\Back");

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(new[] { source });

        Assert.Equal(@"Q1 \> Q2\\Back", Assert.Single(result.Segments).Context);
    }

    [Fact]
    public void Chunk_InfersSourceIdentityWithoutDroppingEnvelopeTitle() {
        ReaderChunk firstChunk = CreateChunk("same", "body");
        firstChunk.SourceId = null;
        firstChunk.Location.Path = "first.txt";
        ReaderChunk secondChunk = CreateChunk("same", "body");
        secondChunk.SourceId = null;
        secondChunk.Location.Path = "second.txt";
        var first = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Title = "Shared title" },
            Chunks = new[] { firstChunk }
        };
        var second = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Title = "Shared title" },
            Chunks = new[] { secondChunk }
        };

        ReaderChunkHierarchyResult firstResult = ReaderHierarchicalChunker.Chunk(first);
        ReaderChunkHierarchyResult secondResult = ReaderHierarchicalChunker.Chunk(second);

        Assert.Equal("Shared title", firstResult.Source.Title);
        Assert.Equal("first.txt", firstResult.Source.Path);
        Assert.NotEqual(firstResult.RootNodeId, secondResult.RootNodeId);
    }

    [Fact]
    public void Chunk_UsesSlideContainerKindForBlockOnlyResults() {
        var document = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "slides.pptx" },
            Pages = new[] {
                new OfficeDocumentPage {
                    Number = 3,
                    Location = new ReaderLocation { SourceBlockKind = "slide" },
                    Blocks = new[] { new OfficeDocumentBlock { Id = "slide-body", Kind = "paragraph", Text = "Slide body" } }
                }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document);

        ReaderLocation location = Assert.Single(result.Chunks).Location;
        Assert.Equal(3, location.Slide);
        Assert.Null(location.Page);
        Assert.Equal("Slide 3", Assert.Single(result.Nodes, node => node.Kind == ReaderChunkHierarchyNodeKind.Container).Title);
    }

    [Fact]
    public void Chunk_PropagatesHeadingSlugsToFallbackBodyBlocks() {
        var document = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { SourceId = "repeated-headings" },
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "overview-a", Kind = "heading", Level = 1, Text = "Overview" },
                new OfficeDocumentBlock { Id = "body-a", Kind = "paragraph", Text = "First body" },
                new OfficeDocumentBlock { Id = "overview-b", Kind = "heading", Level = 1, Text = "Overview" },
                new OfficeDocumentBlock { Id = "body-b", Kind = "paragraph", Text = "Second body" }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document);

        Assert.Equal("overview-a", result.Chunks[1].Location.HeadingSlug);
        Assert.Equal("overview-b", result.Chunks[3].Location.HeadingSlug);
        ReaderChunkHierarchyNode[] headings = result.Nodes
            .Where(node => node.Kind == ReaderChunkHierarchyNodeKind.Heading && node.Title == "Overview")
            .ToArray();
        Assert.Equal(2, headings.Length);
        Assert.All(headings, heading => Assert.Equal(2, heading.ChildNodeIds.Count));
    }

    [Fact]
    public void Chunk_SeparatesRepeatedFallbackAncestorHeadingsByPerLevelSlug() {
        var document = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { SourceId = "repeated-ancestors" },
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "overview-a", Kind = "heading", Level = 1, Text = "Overview" },
                new OfficeDocumentBlock { Id = "details-a", Kind = "heading", Level = 2, Text = "Details" },
                new OfficeDocumentBlock { Id = "body-a", Kind = "paragraph", Text = "First body" },
                new OfficeDocumentBlock { Id = "overview-b", Kind = "heading", Level = 1, Text = "Overview" },
                new OfficeDocumentBlock { Id = "details-b", Kind = "heading", Level = 2, Text = "Details" },
                new OfficeDocumentBlock { Id = "body-b", Kind = "paragraph", Text = "Second body" }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(document);

        ReaderChunkHierarchyNode[] overviews = result.Nodes
            .Where(node => node.Kind == ReaderChunkHierarchyNodeKind.Heading && node.Title == "Overview")
            .ToArray();
        Assert.Equal(2, overviews.Length);
        Assert.All(overviews, overview => Assert.Single(
            overview.ChildNodeIds.Select(id => result.Nodes.Single(node => node.Id == id)),
            child => child.Kind == ReaderChunkHierarchyNodeKind.Heading && child.Title == "Details"));
    }

    [Fact]
    public void Chunk_PreservesFallbackSlugIdentityWhenHeadingDepthCollapses() {
        var document = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { SourceId = "collapsed-repeated-ancestors" },
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "overview-a", Kind = "heading", Level = 1, Text = "Overview" },
                new OfficeDocumentBlock { Id = "details-a", Kind = "heading", Level = 2, Text = "Details" },
                new OfficeDocumentBlock { Id = "body-a", Kind = "paragraph", Text = "First body" },
                new OfficeDocumentBlock { Id = "overview-b", Kind = "heading", Level = 1, Text = "Overview" },
                new OfficeDocumentBlock { Id = "details-b", Kind = "heading", Level = 2, Text = "Details" },
                new OfficeDocumentBlock { Id = "body-b", Kind = "paragraph", Text = "Second body" }
            }
        };

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
            document,
            new ReaderHierarchicalChunkingOptions { MaxHierarchyDepth = 1 });

        ReaderChunkHierarchyNode[] headings = result.Nodes
            .Where(node => node.Kind == ReaderChunkHierarchyNodeKind.Heading)
            .ToArray();
        Assert.Equal(4, headings.Length);
        Assert.Equal(4, headings.Select(heading => heading.Id).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void Chunk_SeparatesEncodedHeadingPathsAfterDepthCollapse() {
        ReaderChunk literal = CreateChunk("literal", "One", headingPath: "A > B > C");
        ReaderHeadingPath.SetHierarchyPath(literal.Location, ReaderHeadingPath.Combine(new[] { "A > B", "C" }));
        ReaderChunk nested = CreateChunk("nested", "Two", headingPath: "A > B > C");
        ReaderHeadingPath.SetHierarchyPath(nested.Location, ReaderHeadingPath.Combine(new[] { "A", "B", "C" }));

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
            new[] { literal, nested },
            new ReaderHierarchicalChunkingOptions { MaxHierarchyDepth = 1 });

        ReaderChunkHierarchyNode[] headings = result.Nodes
            .Where(node => node.Kind == ReaderChunkHierarchyNodeKind.Heading)
            .ToArray();
        Assert.Equal(2, headings.Length);
        Assert.All(headings, heading => Assert.Equal("A > B > C", heading.Title));
        Assert.NotEqual(headings[0].Id, headings[1].Id);
    }

    [Fact]
    public void Chunk_InheritsEnvelopeSourceIdForPathOnlyChunks() {
        ReaderChunk first = CreateChunk("chunk", "body");
        first.SourceId = null;
        first.Location.Path = "shared-name.txt";
        var firstDocument = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { SourceId = "tenant-a", Path = "shared-name.txt" },
            Chunks = new[] { first }
        };
        ReaderChunk second = CreateChunk("chunk", "body");
        second.SourceId = null;
        second.Location.Path = "shared-name.txt";
        var secondDocument = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { SourceId = "tenant-b", Path = "shared-name.txt" },
            Chunks = new[] { second }
        };

        ReaderChunkHierarchyResult firstResult = ReaderHierarchicalChunker.Chunk(firstDocument);
        ReaderChunkHierarchyResult secondResult = ReaderHierarchicalChunker.Chunk(secondDocument);

        Assert.Equal("tenant-a", Assert.Single(firstResult.Chunks).SourceId);
        Assert.NotEqual(Assert.Single(firstResult.Chunks).Id, Assert.Single(secondResult.Chunks).Id);
    }

    [Fact]
    public void Chunk_LengthPrefixesSegmentIdentityFields() {
        ReaderChunk first = CreateChunk("c", "body");
        first.SourceId = "a|b";
        ReaderChunk second = CreateChunk("b|c", "body");
        second.SourceId = "a";

        string firstId = Assert.Single(ReaderHierarchicalChunker.Chunk(new[] { first }).Chunks).Id;
        string secondId = Assert.Single(ReaderHierarchicalChunker.Chunk(new[] { second }).Chunks).Id;

        Assert.NotEqual(firstId, secondId);
    }

    [Fact]
    public void Chunk_InfersSourceFromFirstIdentifiedChunk() {
        ReaderChunk anonymous = CreateChunk("anonymous", "preface");
        anonymous.SourceId = null;
        anonymous.Location.Path = null;
        ReaderChunk identified = CreateChunk("identified", "body");
        identified.SourceId = null;
        identified.Location.Path = "guide.txt";

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(new[] { anonymous, identified });

        Assert.Equal("guide.txt", result.Source.Path);
        Assert.Equal("guide.txt", result.Nodes.Single(node => node.Kind == ReaderChunkHierarchyNodeKind.Document).Title);
    }

    [Fact]
    public void Chunk_DerivesContextTokensFromCombinedTokenizerOutput() {
        ReaderChunk source = CreateChunk("source", "a", page: 3);

        ReaderChunkHierarchyResult result = ReaderHierarchicalChunker.Chunk(
            new[] { source },
            new ReaderHierarchicalChunkingOptions {
                MaxTokens = 10,
                OverlapTokens = 0,
                IncludeContextInText = true,
                TokenCounter = new NonAdditiveContextTokenCounter()
            });

        ReaderChunkSegment segment = Assert.Single(result.Segments);
        Assert.Equal(3, segment.TokenCount);
        Assert.Equal(1, segment.ContentTokenCount);
        Assert.Equal(2, segment.ContextTokenCount);
        Assert.Equal(2, result.ContextTokenCount);
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

        ReaderChunkHierarchyResult staticSync = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadHierarchical(markdown, "note.md", chunkingOptions: options);
        ReaderChunkHierarchyResult staticAsync = await OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadHierarchicalAsync(markdown, "note.md", chunkingOptions: options);
        Assert.Equal(staticSync.ToJson(), staticAsync.ToJson());

        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddMarkdownHandler()
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

    private sealed class NonAdditiveContextTokenCounter : IReaderTokenCounter {
        public string Id => "tests.non-additive-context-v1";

        public int CountTokens(string text) {
            if (text.Length == 0) return 0;
            if (text.EndsWith("\n\n", StringComparison.Ordinal)) return 1;
            return text.Contains("\n\n", StringComparison.Ordinal) ? 3 : 1;
        }
    }

    private sealed class CountingBlockList : IReadOnlyList<OfficeDocumentBlock> {
        private readonly OfficeDocumentBlock _block;

        internal CountingBlockList(OfficeDocumentBlock block, int count) {
            _block = block;
            Count = count;
        }

        public int Count { get; }
        internal int ReadCount { get; private set; }

        public OfficeDocumentBlock this[int index] {
            get {
                if (index < 0 || index >= Count) throw new ArgumentOutOfRangeException(nameof(index));
                ReadCount++;
                return _block;
            }
        }

        public IEnumerator<OfficeDocumentBlock> GetEnumerator() {
            for (int index = 0; index < Count; index++) yield return this[index];
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class CountingPageList : IReadOnlyList<OfficeDocumentPage> {
        private readonly OfficeDocumentPage _page;

        internal CountingPageList(OfficeDocumentPage page, int count) {
            _page = page;
            Count = count;
        }

        public int Count { get; }
        internal int ReadCount { get; private set; }

        public OfficeDocumentPage this[int index] {
            get {
                if (index < 0 || index >= Count) throw new ArgumentOutOfRangeException(nameof(index));
                ReadCount++;
                return _page;
            }
        }

        public IEnumerator<OfficeDocumentPage> GetEnumerator() {
            for (int index = 0; index < Count; index++) yield return this[index];
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

}

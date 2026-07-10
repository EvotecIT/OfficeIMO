using OfficeIMO.Reader;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderProcessorPipelineTests {
    [Fact]
    public void PipelineBuilder_FreezesOrderAndRejectsDuplicateIds() {
        var builder = new OfficeDocumentProcessorPipelineBuilder()
            .Add(PassThrough("first"));
        OfficeDocumentProcessorPipeline firstSnapshot = builder.Build();

        builder.Add(PassThrough("second"));
        OfficeDocumentProcessorPipeline secondSnapshot = builder.Build();

        Assert.Equal(new[] { "first" }, firstSnapshot.Processors.Select(processor => processor.Id));
        Assert.Equal(new[] { "first", "second" }, secondSnapshot.Processors.Select(processor => processor.Id));
        Assert.Throws<InvalidOperationException>(() => builder.Add(PassThrough("FIRST")));
        Assert.Throws<ArgumentException>(() => builder.Add(new InvalidIdProcessor()));
    }

    [Fact]
    public void Pipeline_ExecutesSynchronouslyInRegistrationOrder() {
        var observed = new List<string>();
        OfficeDocumentProcessorPipeline pipeline = new OfficeDocumentProcessorPipelineBuilder()
            .Add(Recording("first", observed))
            .Add(Recording("second", observed))
            .Build();

        OfficeDocumentProcessingResult result = pipeline.Process(new OfficeDocumentReadResult());

        Assert.Equal(new[] { "first:0/2", "second:1/2" }, observed);
        Assert.True(result.Succeeded);
        Assert.Equal(new[] {
            OfficeDocumentProcessorStepStatus.Completed,
            OfficeDocumentProcessorStepStatus.Completed
        }, result.Steps.Select(step => step.Status));
    }

    [Fact]
    public async Task Pipeline_UsesExplicitAsyncProcessorsInOrder() {
        var observed = new List<string>();
        var processor = new DelegateOfficeDocumentProcessor(
            "async",
            (document, _) => throw new InvalidOperationException("Sync path should not run."),
            async (document, context) => {
                await Task.Yield();
                observed.Add(context.ProcessorId);
                document.Source.Title = "processed asynchronously";
                return document;
            });
        OfficeDocumentProcessorPipeline pipeline = new OfficeDocumentProcessorPipelineBuilder().Add(processor).Build();

        OfficeDocumentProcessingResult result = await pipeline.ProcessAsync(new OfficeDocumentReadResult());

        Assert.Equal("async", Assert.Single(observed));
        Assert.Equal("processed asynchronously", result.Document.Source.Title);
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void Pipeline_ContinuePolicyAddsDiagnosticAndRunsLaterSteps() {
        bool laterRan = false;
        OfficeDocumentProcessorPipeline pipeline = new OfficeDocumentProcessorPipelineBuilder()
            .Add(Failing("broken"))
            .Add(new DelegateOfficeDocumentProcessor("later", (document, _) => {
                laterRan = true;
                return document;
            }))
            .Build();

        OfficeDocumentProcessingResult result = pipeline.Process(
            new OfficeDocumentReadResult(),
            new OfficeDocumentProcessingOptions {
                FailureBehavior = OfficeDocumentProcessorFailureBehavior.ContinueWithDiagnostic
            });

        Assert.True(laterRan);
        Assert.False(result.Succeeded);
        Assert.Equal(OfficeDocumentProcessorStepStatus.Failed, result.Steps[0].Status);
        Assert.Equal(OfficeDocumentProcessorStepStatus.Completed, result.Steps[1].Status);
        OfficeDocumentDiagnostic diagnostic = Assert.Single(result.Document.Diagnostics);
        Assert.Equal("processor-failed", diagnostic.Code);
        Assert.Equal("broken", diagnostic.Attributes["processorId"]);
        Assert.True(diagnostic.IsRecoverable == true);
    }

    [Fact]
    public void Pipeline_StopPolicyMarksLaterStepsSkipped() {
        bool laterRan = false;
        OfficeDocumentProcessorPipeline pipeline = new OfficeDocumentProcessorPipelineBuilder()
            .Add(Failing("broken"))
            .Add(new DelegateOfficeDocumentProcessor("later", (document, _) => {
                laterRan = true;
                return document;
            }))
            .Build();

        OfficeDocumentProcessingResult result = pipeline.Process(
            new OfficeDocumentReadResult(),
            new OfficeDocumentProcessingOptions {
                FailureBehavior = OfficeDocumentProcessorFailureBehavior.StopWithDiagnostic
            });

        Assert.False(laterRan);
        Assert.Equal(OfficeDocumentProcessorStepStatus.Failed, result.Steps[0].Status);
        Assert.Equal(OfficeDocumentProcessorStepStatus.Skipped, result.Steps[1].Status);
        Assert.True(Assert.Single(result.Document.Diagnostics).IsRecoverable == false);
    }

    [Fact]
    public void Pipeline_ThrowPolicyWrapsProcessorIdentity() {
        OfficeDocumentProcessorPipeline pipeline = new OfficeDocumentProcessorPipelineBuilder()
            .Add(Failing("broken"))
            .Build();

        OfficeDocumentProcessorException exception = Assert.Throws<OfficeDocumentProcessorException>(
            () => pipeline.Process(new OfficeDocumentReadResult()));

        Assert.Equal("broken", exception.ProcessorId);
        Assert.Equal(0, exception.ProcessorIndex);
        Assert.IsType<FormatException>(exception.InnerException);
    }

    [Fact]
    public async Task InstanceReader_AppliesFrozenPipelineToEveryDocumentSurface() {
        byte[] markdown = Encoding.UTF8.GetBytes("# Title\n\nBody");
        var builder = new OfficeDocumentReaderBuilder()
            .AddProcessor(new DelegateOfficeDocumentProcessor("mark", (document, _) => {
                foreach (ReaderChunk chunk in document.Chunks) chunk.Text = "processed:" + chunk.Text;
                document.Metadata = new[] {
                    new OfficeDocumentMetadataEntry { Id = "processor", Category = "processor", Name = "Applied", Value = "true" }
                };
                return document;
            }));
        OfficeDocumentReader first = builder.Build();
        builder.AddProcessor(PassThrough("later"));
        OfficeDocumentReader second = builder.Build();

        Assert.Equal(1, first.ProcessorPipeline.Count);
        Assert.Equal(2, second.ProcessorPipeline.Count);
        Assert.StartsWith("processed:", Assert.Single(first.Read(markdown, "note.md")).Text, StringComparison.Ordinal);
        Assert.StartsWith("processed:", Assert.Single(first.ReadDocument(markdown, "note.md").Chunks).Text, StringComparison.Ordinal);
        Assert.StartsWith("processed:", Assert.Single(await first.ReadAsync(markdown, "note.md")).Text, StringComparison.Ordinal);
        Assert.StartsWith("processed:", Assert.Single((await first.ReadDocumentAsync(markdown, "note.md")).Chunks).Text, StringComparison.Ordinal);

        using JsonDocument json = JsonDocument.Parse(first.ReadDocumentJson(markdown, "note.md"));
        Assert.Equal("Applied", json.RootElement.GetProperty("metadata")[0].GetProperty("name").GetString());
    }

    [Fact]
    public void InstanceReader_ProjectsNullProcessorChunksAsAnEmptyCollection() {
        byte[] markdown = Encoding.UTF8.GetBytes("Body");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddProcessor(new DelegateOfficeDocumentProcessor("clear-chunks", (document, _) => {
                document.Chunks = null!;
                return document;
            }))
            .Build();

        Assert.Empty(reader.Read(markdown, "note.md"));
    }

    [Fact]
    public void BuiltInProcessorsNormalizeAndFilterSharedModels() {
        var kept = new OfficeDocumentAsset { Id = "keep", Kind = "image", LengthBytes = 10 };
        var removed = new OfficeDocumentAsset { Id = "drop", Kind = "preview", LengthBytes = 1000 };
        var table = new ReaderTable {
            Title = "  Settings  ",
            Columns = new[] { " Key ", " " },
            Rows = new[] { (IReadOnlyList<string>)new[] { " A ", " 1 ", " extra " } }
        };
        var link = new OfficeDocumentLink { Id = " l1 ", Kind = " URI ", Uri = " https://example.test " };
        var result = new OfficeDocumentReadResult {
            Blocks = new[] { new OfficeDocumentBlock { Id = "h", Kind = " Heading ", Text = " Title ", Level = 9 } },
            Tables = new[] { table },
            Links = new[] { link },
            Assets = new[] { kept, removed },
            OcrCandidates = new[] {
                new OfficeDocumentOcrCandidate { Id = "keep-ocr", AssetId = "keep" },
                new OfficeDocumentOcrCandidate { Id = "drop-ocr", AssetId = "drop" }
            }
        };
        OfficeDocumentProcessorPipeline pipeline = new OfficeDocumentProcessorPipelineBuilder()
            .Add(new OfficeDocumentBlockNormalizationProcessor())
            .Add(new OfficeDocumentTableNormalizationProcessor())
            .Add(new OfficeDocumentLinkNormalizationProcessor())
            .Add(new OfficeDocumentAssetFilterProcessor(asset => asset.LengthBytes <= 100))
            .Build();

        pipeline.Process(result);

        Assert.Equal("heading", Assert.Single(result.Blocks).Kind);
        Assert.Equal("Title", Assert.Single(result.Blocks).Text);
        Assert.Equal(6, Assert.Single(result.Blocks).Level);
        Assert.Equal(new[] { "Key", "Column2", "Column3" }, table.Columns);
        Assert.Equal(new[] { "A", "1", "extra" }, Assert.Single(table.Rows));
        Assert.Equal("Settings", table.Title);
        Assert.Equal("uri", link.Kind);
        Assert.Equal("https://example.test", link.Uri);
        Assert.Equal("keep", Assert.Single(result.Assets).Id);
        Assert.Equal("keep-ocr", Assert.Single(result.OcrCandidates).Id);
    }

    [Fact]
    public void ArtifactClassifierUsesRepeatedPageBoundaries() {
        OfficeDocumentPage[] pages = Enumerable.Range(1, 3)
            .Select(page => new OfficeDocumentPage {
                Number = page,
                Blocks = new[] {
                    new OfficeDocumentBlock { Id = "h" + page, Kind = "paragraph", Text = "Company Confidential" },
                    new OfficeDocumentBlock { Id = "b" + page, Kind = "paragraph", Text = "Body " + page },
                    new OfficeDocumentBlock { Id = "f" + page, Kind = "paragraph", Text = "Page footer" }
                }
            })
            .ToArray();
        var document = new OfficeDocumentReadResult { Pages = pages };

        new OfficeDocumentProcessorPipelineBuilder()
            .Add(new OfficeDocumentArtifactClassificationProcessor())
            .Build()
            .Process(document);

        Assert.All(pages, page => Assert.Equal("header", page.Blocks[0].Kind));
        Assert.All(pages, page => Assert.Equal("footer", page.Blocks[2].Kind));
        Assert.All(pages, page => Assert.Equal("paragraph", page.Blocks[1].Kind));
    }

    [Fact]
    public void ArtifactClassifierDoesNotRelabelMatchingBodyText() {
        OfficeDocumentPage[] pages = Enumerable.Range(1, 3)
            .Select(page => new OfficeDocumentPage {
                Number = page,
                Blocks = new[] {
                    new OfficeDocumentBlock { Id = "h" + page, Kind = "paragraph", Text = "Repeated text" },
                    new OfficeDocumentBlock { Id = "b" + page, Kind = "paragraph", Text = "Repeated text" },
                    new OfficeDocumentBlock { Id = "f" + page, Kind = "paragraph", Text = "Footer " + page }
                }
            })
            .ToArray();
        var document = new OfficeDocumentReadResult { Pages = pages };

        new OfficeDocumentProcessorPipelineBuilder()
            .Add(new OfficeDocumentArtifactClassificationProcessor(
                new OfficeDocumentArtifactClassificationOptions { BoundaryBlockCount = 1 }))
            .Build()
            .Process(document);

        Assert.All(pages, page => Assert.Equal("header", page.Blocks[0].Kind));
        Assert.All(pages, page => Assert.Equal("paragraph", page.Blocks[1].Kind));
        Assert.All(pages, page => Assert.Equal("paragraph", page.Blocks[2].Kind));
    }

    private static IOfficeDocumentProcessor PassThrough(string id) =>
        new DelegateOfficeDocumentProcessor(id, (document, _) => document);

    private static IOfficeDocumentProcessor Recording(string id, ICollection<string> observed) =>
        new DelegateOfficeDocumentProcessor(id, (document, context) => {
            observed.Add($"{context.ProcessorId}:{context.ProcessorIndex}/{context.ProcessorCount}");
            return document;
        });

    private static IOfficeDocumentProcessor Failing(string id) =>
        new DelegateOfficeDocumentProcessor(id, (_, _) => throw new FormatException("bad input"));

    private sealed class InvalidIdProcessor : IOfficeDocumentProcessor {
        public string Id => " invalid ";

        public OfficeDocumentReadResult Process(
            OfficeDocumentReadResult document,
            OfficeDocumentProcessorContext context) => document;

        public Task<OfficeDocumentReadResult> ProcessAsync(
            OfficeDocumentReadResult document,
            OfficeDocumentProcessorContext context) => Task.FromResult(document);
    }
}

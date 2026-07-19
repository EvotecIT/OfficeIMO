using OfficeIMO.Reader;
using OfficeIMO.Reader.Markdown;
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
        var processor = new DelegateAsyncOfficeDocumentProcessor(
            "async",
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
    public void Pipeline_ContinuePolicyRetainsFailureDiagnosticWhenLaterStepReplacesDocument() {
        OfficeDocumentProcessorPipeline pipeline = new OfficeDocumentProcessorPipelineBuilder()
            .Add(Failing("broken"))
            .Add(new DelegateOfficeDocumentProcessor("replace", (_, _) => new OfficeDocumentReadResult()))
            .Build();

        OfficeDocumentProcessingResult result = pipeline.Process(
            new OfficeDocumentReadResult(),
            new OfficeDocumentProcessingOptions {
                FailureBehavior = OfficeDocumentProcessorFailureBehavior.ContinueWithDiagnostic
            });

        OfficeDocumentDiagnostic diagnostic = Assert.Single(result.Document.Diagnostics);
        Assert.Equal("processor-failed", diagnostic.Code);
        Assert.Equal("broken", diagnostic.Attributes["processorId"]);
    }

    [Fact]
    public async Task Pipeline_ContinuePolicyRetainsFailureDiagnosticWhenAsyncStepReplacesDocument() {
        OfficeDocumentProcessorPipeline pipeline = new OfficeDocumentProcessorPipelineBuilder()
            .Add(Failing("broken"))
            .Add(new DelegateOfficeDocumentProcessor("replace", (_, _) => new OfficeDocumentReadResult()))
            .Build();

        OfficeDocumentProcessingResult result = await pipeline.ProcessAsync(
            new OfficeDocumentReadResult(),
            new OfficeDocumentProcessingOptions {
                FailureBehavior = OfficeDocumentProcessorFailureBehavior.ContinueWithDiagnostic
            });

        OfficeDocumentDiagnostic diagnostic = Assert.Single(result.Document.Diagnostics);
        Assert.Equal("processor-failed", diagnostic.Code);
        Assert.Equal("broken", diagnostic.Attributes["processorId"]);
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
            .AddMarkdownHandler()
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
    public async Task InstanceReader_RefreshesProcessorReplacementChunkMetadata() {
        const string processedText = "processor replacement body";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddMarkdownHandler()
            .AddProcessor(new DelegateOfficeDocumentProcessor("replace-chunk", (document, _) => {
                return new OfficeDocumentReadResult {
                    Chunks = new[] {
                        new ReaderChunk {
                            Id = "replacement",
                            Kind = ReaderInputKind.Markdown,
                            Text = processedText,
                            SourceId = "stale-source",
                            SourceHash = "stale-source-hash",
                            ChunkHash = "stale-chunk-hash",
                            TokenEstimate = 999
                        }
                    }
                };
            }))
            .Build();
        byte[] markdown = Encoding.UTF8.GetBytes("Original body");

        OfficeDocumentReadResult sync = reader.ReadDocument(markdown, "note.md");
        OfficeDocumentReadResult asyncResult = await reader.ReadDocumentAsync(markdown, "note.md");

        AssertRefreshedChunk(sync, processedText);
        AssertRefreshedChunk(asyncResult, processedText);
    }

    [Fact]
    public void InstanceReader_ProjectsNullProcessorChunksAsAnEmptyCollection() {
        byte[] markdown = Encoding.UTF8.GetBytes("Body");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddMarkdownHandler()
            .AddProcessor(new DelegateOfficeDocumentProcessor("clear-chunks", (document, _) => {
                document.Chunks = null!;
                return document;
            }))
            .Build();

        Assert.Empty(reader.Read(markdown, "note.md"));
    }

    [Fact]
    public void InstanceReader_PreservesProcessorChunkHashWhenHashComputationIsDisabled() {
        const string suppliedHash = "processor-owned-hash";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddMarkdownHandler()
            .AddProcessor(new DelegateOfficeDocumentProcessor("supply-hash", (document, _) => {
                document.Chunks = new[] {
                    new ReaderChunk { Id = "replacement", Text = "processed", ChunkHash = suppliedHash }
                };
                return document;
            }))
            .Build();

        ReaderChunk chunk = Assert.Single(reader.Read(
            Encoding.UTF8.GetBytes("Original body"),
            "note.md",
            new ReaderOptions { ComputeHashes = false }));

        Assert.Equal(suppliedHash, chunk.ChunkHash);
    }

    [Fact]
    public void InstanceReader_RebuildsChunkDerivedAggregatesAfterTextProcessing() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddMarkdownHandler()
            .AddProcessor(new DelegateOfficeDocumentProcessor("redact", (document, _) => {
                foreach (ReaderChunk chunk in document.Chunks) chunk.Text = "[redacted]";
                return document;
            }))
            .Build();
        byte[] source = Encoding.UTF8.GetBytes("# Secret\n\nSensitive body");

        OfficeDocumentReadResult document = reader.ReadDocument(source, "secret.md");

        Assert.Equal("[redacted]", document.Markdown);
        Assert.All(document.Blocks, block => Assert.Equal("[redacted]", block.Text));
        string json = reader.ReadDocumentJson(source, "secret.md");
        Assert.Contains("[redacted]", json, StringComparison.Ordinal);
        Assert.DoesNotContain("Sensitive body", json, StringComparison.Ordinal);
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
            Metadata = new[] { new OfficeDocumentMetadataEntry { Id = "reader-asset-count", Value = "2", ValueType = "count" } },
            OcrCandidates = new[] {
                new OfficeDocumentOcrCandidate { Id = "keep-ocr", AssetId = "keep", Reason = "Keep OCR", Location = new ReaderLocation { BlockAnchor = "keep" } },
                new OfficeDocumentOcrCandidate { Id = "drop-ocr", AssetId = "drop", Reason = "Drop OCR", Location = new ReaderLocation { BlockAnchor = "drop" } }
            },
            Diagnostics = new[] {
                new OfficeDocumentDiagnostic { Code = "ocr-needed", Message = "Keep OCR", Location = new ReaderLocation { BlockAnchor = "keep" } },
                new OfficeDocumentDiagnostic { Code = "ocr-needed", Message = "Drop OCR", Location = new ReaderLocation { BlockAnchor = "drop" } }
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
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "reader-asset-count").Value);
        Assert.Equal("keep-ocr", Assert.Single(result.OcrCandidates).Id);
        OfficeDocumentDiagnostic ocrDiagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "ocr-needed");
        Assert.Equal("Keep OCR", ocrDiagnostic.Message);
        Assert.Equal("keep", ocrDiagnostic.Location!.BlockAnchor);
    }

    [Fact]
    public void TableNormalization_UpdatesAggregateAndChunkTableInstances() {
        var aggregate = new ReaderTable {
            Columns = new[] { " Key " },
            Rows = new[] { (IReadOnlyList<string>)new[] { " Value ", " Extra " } }
        };
        var chunkTable = new ReaderTable {
            Columns = new[] { " Key " },
            Rows = new[] { (IReadOnlyList<string>)new[] { " Value ", " Extra " } }
        };
        var document = new OfficeDocumentReadResult {
            Tables = new[] { aggregate },
            Chunks = new[] { new ReaderChunk { Tables = new[] { chunkTable } } }
        };

        new OfficeDocumentProcessorPipelineBuilder()
            .Add(new OfficeDocumentTableNormalizationProcessor())
            .Build()
            .Process(document);

        Assert.Equal(new[] { "Key", "Column2" }, aggregate.Columns);
        Assert.Equal(new[] { "Key", "Column2" }, chunkTable.Columns);
        Assert.Equal(new[] { "Value", "Extra" }, Assert.Single(chunkTable.Rows));
    }

    [Fact]
    public void StructuredExtraction_UsesPageLocationFallbackAndDeduplicatesPageTableClone() {
        var aggregate = new ReaderTable {
            Title = "Settings",
            Location = new ReaderLocation { Path = "settings.pdf", Page = 7, TableIndex = 0 },
            Columns = new[] { "Key", "Value" },
            Rows = new[] { (IReadOnlyList<string>)new[] { "Mode", "Safe" } }
        };
        var pageTable = new ReaderTable {
            Title = aggregate.Title,
            Columns = aggregate.Columns,
            Rows = aggregate.Rows
        };
        var document = new OfficeDocumentReadResult {
            Tables = new[] { aggregate },
            Pages = new[] {
                new OfficeDocumentPage {
                    Number = 7,
                    Location = new ReaderLocation { Path = "settings.pdf" },
                    Tables = new[] { pageTable }
                }
            }
        };

        OfficeDocumentStructuredExtractionResult extracted = OfficeDocumentStructuredExtractor.Extract(document);

        Assert.Single(extracted.Records, record => record.Category == "key-value");
        ReaderTable selected = Assert.Single(extracted.Tables);
        Assert.Equal(7, selected.Location!.Page);
        Assert.Equal("settings.pdf", selected.Location.Path);
    }

    [Fact]
    public void AssetFilter_KeepsOcrForAnAssetIdThatSurvivesInAnotherScope() {
        var document = new OfficeDocumentReadResult {
            Assets = new[] {
                new OfficeDocumentAsset { Id = "shared", LengthBytes = 10 }
            },
            Pages = new[] {
                new OfficeDocumentPage {
                    Assets = new[] { new OfficeDocumentAsset { Id = "shared", LengthBytes = null } }
                }
            },
            OcrCandidates = new[] {
                new OfficeDocumentOcrCandidate { Id = "shared-ocr", AssetId = "shared" }
            },
            Diagnostics = new[] {
                new OfficeDocumentDiagnostic { Code = "ocr-needed", Message = "OCR shared asset" }
            }
        };

        new OfficeDocumentProcessorPipelineBuilder()
            .Add(new OfficeDocumentAssetFilterProcessor(asset => asset.LengthBytes.HasValue))
            .Build()
            .Process(document);

        Assert.Equal("shared", Assert.Single(document.Assets).Id);
        Assert.Empty(Assert.Single(document.Pages).Assets);
        Assert.Equal("shared-ocr", Assert.Single(document.OcrCandidates).Id);
        Assert.Equal("ocr-needed", Assert.Single(document.Diagnostics).Code);
    }

    [Fact]
    public void AssetFilter_PreservesUnrelatedOcrDiagnosticsWithoutCandidates() {
        var document = new OfficeDocumentReadResult {
            Assets = new[] { new OfficeDocumentAsset { Id = "drop", LengthBytes = 1000 } },
            OcrCandidates = new[] {
                new OfficeDocumentOcrCandidate {
                    Id = "drop-ocr",
                    AssetId = "drop",
                    Reason = "Removed image OCR",
                    Location = new ReaderLocation { BlockAnchor = "drop" }
                }
            },
            Diagnostics = new[] {
                new OfficeDocumentDiagnostic {
                    Code = "ocr-needed",
                    Message = "Removed image OCR",
                    Location = new ReaderLocation { BlockAnchor = "drop" }
                },
                new OfficeDocumentDiagnostic {
                    Code = "ocr-needed",
                    Message = "Page-level OCR remains required",
                    Location = new ReaderLocation { Page = 9 },
                    Attributes = new Dictionary<string, string> { ["providerEvidence"] = "adapter-owned" }
                }
            }
        };

        new OfficeDocumentProcessorPipelineBuilder()
            .Add(new OfficeDocumentAssetFilterProcessor(asset => asset.LengthBytes <= 100))
            .Build()
            .Process(document);

        Assert.Empty(document.Assets);
        Assert.Empty(document.OcrCandidates);
        OfficeDocumentDiagnostic remaining = Assert.Single(document.Diagnostics);
        Assert.Equal("Page-level OCR remains required", remaining.Message);
        Assert.Equal("adapter-owned", remaining.Attributes["providerEvidence"]);
    }

    [Fact]
    public void BlockNormalization_AllowsNullKindWhenKindNormalizationIsDisabled() {
        var block = new OfficeDocumentBlock { Kind = null!, Level = 3 };
        var document = new OfficeDocumentReadResult { Blocks = new[] { block } };
        new OfficeDocumentProcessorPipelineBuilder()
            .Add(new OfficeDocumentBlockNormalizationProcessor(
                new OfficeDocumentBlockNormalizationOptions {
                    NormalizeKinds = false,
                    NormalizeLevels = true
                }))
            .Build()
            .Process(document);

        Assert.Null(block.Kind);
        Assert.Equal(3, block.Level);
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

    private static void AssertRefreshedChunk(OfficeDocumentReadResult document, string expectedText) {
        ReaderChunk chunk = Assert.Single(document.Chunks);
        Assert.Equal(document.Source.SourceId, chunk.SourceId);
        Assert.Equal(document.Source.SourceHash, chunk.SourceHash);
        Assert.Equal((expectedText.Length + 3) / 4, chunk.TokenEstimate);
        Assert.NotEqual("stale-chunk-hash", chunk.ChunkHash);
        Assert.Equal(64, chunk.ChunkHash?.Length);
    }

    private sealed class InvalidIdProcessor : IOfficeDocumentProcessor {
        public string Id => " invalid ";
    }
}

using OfficeIMO.Reader;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrCoreTests {
    [Fact]
    public async Task ApplyOcrAsync_PreservesCandidateOrderAndDetailedSpansUnderConcurrency() {
        OfficeDocumentReadResult source = CreateDocument(2);
        var engine = new RecordingOcrEngine();

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            Language = "en",
            MaxDegreeOfParallelism = 2
        });

        Assert.Equal(2, execution.Report.CandidateCount);
        Assert.Equal(2, execution.Report.AttemptedCandidateCount);
        Assert.Equal(2, execution.Report.RecognizedCandidateCount);
        Assert.Equal(0, execution.Report.SkippedCandidateCount);
        Assert.Equal(2, execution.Report.EffectiveDegreeOfParallelism);
        Assert.Equal(2, execution.Report.LineSpanCount);
        Assert.Equal(2, execution.Report.WordSpanCount);
        Assert.Equal(2, execution.Report.CharacterSpanCount);
        Assert.True(engine.MaximumConcurrentCalls >= 2);
        Assert.Equal(new[] { "ocr-1", "ocr-2" }, execution.Recognitions.Select(item => item.CandidateId).ToArray());
        Assert.Contains(execution.Recognitions[0].Result.Spans, span => span.Level == OfficeOcrTextSpanLevel.Line);
        Assert.Contains(execution.Recognitions[0].Result.Spans, span => span.Level == OfficeOcrTextSpanLevel.Word);
        Assert.Contains(execution.Recognitions[0].Result.Spans, span => span.Level == OfficeOcrTextSpanLevel.Character);
        Assert.Empty(execution.Document.OcrCandidates);
        Assert.Equal(2, execution.Document.Blocks.Count(block => block.Kind == "ocr-text"));
        Assert.DoesNotContain(execution.Document.Diagnostics, diagnostic => diagnostic.Code == "ocr-needed");
        Assert.Contains("officeimo.reader.ocr-execution", execution.Document.CapabilitiesUsed);
        Assert.Contains("officeimo.reader.ocr-engine.fixture-engine", execution.Document.CapabilitiesUsed);
        Assert.Equal("2", Assert.Single(execution.Document.Metadata, item => item.Id == "reader-ocr-execution-recognized-count").Value);
    }

    [Fact]
    public async Task ApplyOcrAsync_EnforcesCandidateAssetHashAndPayloadLimitsBeforeCallingEngine() {
        OfficeDocumentReadResult source = CreateDocument(5);
        source.OcrCandidates[1].AssetId = "missing";
        source.Assets[2].PayloadHash = new string('0', 64);
        source.Assets[3].PayloadBytes = new byte[] { 1, 2, 3 };
        source.Assets[3].LengthBytes = 3;
        var engine = new RecordingOcrEngine();

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxCandidates = 4,
            MaxInputBytesPerCandidate = 2,
            MaxTotalInputBytes = 8
        });

        Assert.Equal(new[] { "ocr-1" }, engine.CandidateIds);
        Assert.Equal(1, execution.Report.AttemptedCandidateCount);
        Assert.Equal(1, execution.Report.RecognizedCandidateCount);
        Assert.Equal(4, execution.Report.SkippedCandidateCount);
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-candidate-limit");
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-asset-missing");
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-payload-hash-mismatch");
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-input-limit");
        Assert.Equal(4, execution.Document.OcrCandidates.Count);
    }

    [Fact]
    public async Task ApplyOcrAsync_DoesNotResolveMultiImagePageToItsFirstImageAsset() {
        OfficeDocumentReadResult source = CreateDocument(2);
        source.OcrCandidates = new[] {
            new OfficeDocumentOcrCandidate {
                Id = "page-ocr",
                Kind = "page",
                AssetId = source.Assets[0].Id,
                ImageCount = 2,
                Location = source.OcrCandidates[0].Location
            }
        };
        var engine = new RecordingOcrEngine();

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine);

        Assert.Empty(engine.CandidateIds);
        Assert.Equal(0, execution.Report.AttemptedCandidateCount);
        Assert.Single(execution.Document.OcrCandidates);
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-asset-ambiguous");
    }

    [Fact]
    public async Task ApplyOcrAsync_RejectsUnknownMediaTypeForRestrictedEngine() {
        OfficeDocumentReadResult source = CreateDocument(1);
        source.Assets[0].MediaType = null;
        var engine = new RecordingOcrEngine();

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine);

        Assert.Empty(engine.CandidateIds);
        Assert.Equal(0, execution.Report.AttemptedCandidateCount);
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-media-type-unsupported");
    }

    [Fact]
    public async Task ApplyOcrAsync_BoundsProviderTextSpansAndConfidenceDiagnostics() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var engine = new DelegateOfficeOcrEngine("bounded-fixture", (request, cancellationToken) => new ValueTask<OfficeOcrEngineResult>(new OfficeOcrEngineResult {
            Text = "1234567890",
            Confidence = 1.5,
            Spans = new[] {
                new OfficeOcrTextSpan { Sequence = 0, Level = OfficeOcrTextSpanLevel.Line, Text = "1234567890", Confidence = -0.5 },
                new OfficeOcrTextSpan { Sequence = 1, Level = OfficeOcrTextSpanLevel.Word, Text = "12345" },
                new OfficeOcrTextSpan { Sequence = 2, Level = OfficeOcrTextSpanLevel.Character, Text = "1" }
            }
        }));

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxRecognizedCharactersPerCandidate = 5,
            MaxSpansPerCandidate = 2
        });

        Assert.Equal("12345", Assert.Single(execution.Document.Blocks, block => block.Kind == "ocr-text").Text);
        OfficeOcrEngineResult result = Assert.Single(execution.Recognitions).Result;
        Assert.Equal(1D, result.Confidence);
        Assert.Equal(0D, result.Spans[0].Confidence);
        Assert.Equal(2, result.Spans.Count);
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-text-limit");
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-span-limit");
        Assert.Single(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-confidence-out-of-range");
    }

    [Fact]
    public async Task ApplyOcrAsync_ConvertsPerCandidateTimeoutToRecoverableDiagnostic() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var engine = new DelegateOfficeOcrEngine("slow-fixture", async (request, cancellationToken) => {
            await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
            return new OfficeOcrEngineResult { Text = "late" };
        });

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            CandidateTimeout = TimeSpan.FromMilliseconds(20),
            ContinueOnError = true
        });

        Assert.Equal(1, execution.Report.FailedCandidateCount);
        Assert.Equal(0, execution.Report.RecognizedCandidateCount);
        Assert.Empty(execution.Recognitions);
        Assert.Single(execution.Document.OcrCandidates);
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout" && diagnostic.IsRecoverable == true);
    }

    [Fact]
    public async Task ApplyOcrAsync_ArmsTimeoutBeforeInvokingSynchronousProviderWork() {
        OfficeDocumentReadResult source = CreateDocument(1);
        bool observedCancellation = false;
        var engine = new DelegateOfficeOcrEngine("synchronous-fixture", (_, cancellationToken) => {
            DateTimeOffset deadline = DateTimeOffset.UtcNow.AddSeconds(2);
            while (!cancellationToken.IsCancellationRequested && DateTimeOffset.UtcNow < deadline) {
                Thread.SpinWait(1000);
            }
            observedCancellation = cancellationToken.IsCancellationRequested;
            cancellationToken.ThrowIfCancellationRequested();
            return new ValueTask<OfficeOcrEngineResult>(new OfficeOcrEngineResult { Text = "late" });
        });

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            CandidateTimeout = TimeSpan.FromMilliseconds(20),
            ContinueOnError = true,
        });

        Assert.True(observedCancellation);
        Assert.Equal(1, execution.Report.FailedCandidateCount);
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout");
    }

    [Fact]
    public async Task ApplyOcrAsync_EnforcesTimeoutWhenSynchronousEngineIgnoresCancellation() {
        OfficeDocumentReadResult source = CreateDocument(1);
        using var releaseProvider = new ManualResetEventSlim(false);
        var engine = new DelegateOfficeOcrEngine("synchronous-non-cooperative-fixture", (_, _) => {
            releaseProvider.Wait();
            return new ValueTask<OfficeOcrEngineResult>(new OfficeOcrEngineResult { Text = "late" });
        });

        try {
            Task<OfficeDocumentOcrExecutionResult> executionTask = source.ApplyOcrAsync(engine,
                new OfficeDocumentOcrExecutionOptions {
                    CandidateTimeout = TimeSpan.FromMilliseconds(20),
                    ContinueOnError = true,
                });
            Task completed = await Task.WhenAny(executionTask, Task.Delay(TimeSpan.FromSeconds(2)));

            Assert.Same(executionTask, completed);
            OfficeDocumentOcrExecutionResult execution = await executionTask;
            Assert.Equal(1, execution.Report.FailedCandidateCount);
            Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout");
        } finally {
            releaseProvider.Set();
        }
    }

    [Fact]
    public async Task ApplyOcrAsync_EnforcesTimeoutWhenEngineIgnoresCancellation() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var completion = new TaskCompletionSource<OfficeOcrEngineResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        var engine = new DelegateOfficeOcrEngine(
            "non-cooperative-fixture",
            (_, _) => new ValueTask<OfficeOcrEngineResult>(completion.Task));

        try {
            Task<OfficeDocumentOcrExecutionResult> executionTask = source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
                CandidateTimeout = TimeSpan.FromMilliseconds(20),
                ContinueOnError = true
            });
            Task completed = await Task.WhenAny(executionTask, Task.Delay(TimeSpan.FromSeconds(2)));

            Assert.Same(executionTask, completed);
            OfficeDocumentOcrExecutionResult execution = await executionTask;
            Assert.Equal(1, execution.Report.FailedCandidateCount);
            Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout");
        } finally {
            completion.TrySetResult(new OfficeOcrEngineResult { Text = "late" });
        }
    }

    [Fact]
    public async Task ApplyOcrAsync_RemovesNonFiniteConfidenceAndNullProviderDiagnostics() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var engine = new DelegateOfficeOcrEngine("permissive-fixture", (_, _) => new ValueTask<OfficeOcrEngineResult>(new OfficeOcrEngineResult {
            Text = "recognized",
            Confidence = double.NaN,
            Spans = new[] {
                new OfficeOcrTextSpan { Sequence = 0, Level = OfficeOcrTextSpanLevel.Word, Text = "recognized", Confidence = double.PositiveInfinity }
            },
            Diagnostics = new OfficeDocumentDiagnostic[] {
                null!,
                new OfficeDocumentDiagnostic { Code = "provider-warning", Message = "Provider warning." }
            }
        }));

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine);

        OfficeOcrEngineResult result = Assert.Single(execution.Recognitions).Result;
        Assert.Null(result.Confidence);
        Assert.Null(Assert.Single(result.Spans).Confidence);
        OfficeDocumentDiagnostic providerDiagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(OfficeDocumentDiagnosticCategory.Ocr, providerDiagnostic.Category);
        Assert.Equal("permissive-fixture", providerDiagnostic.Source);
        Assert.NotNull(providerDiagnostic.Location);
        Assert.Single(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-confidence-out-of-range");
    }

    [Fact]
    public async Task ApplyOcrAsync_SerializesConcurrentExecutionsForNonConcurrentEngineInstance() {
        var engine = new RecordingOcrEngine(supportsConcurrentRequests: false);

        await Task.WhenAll(
            CreateDocument(1).ApplyOcrAsync(engine),
            CreateDocument(1).ApplyOcrAsync(engine));

        Assert.Equal(1, engine.MaximumConcurrentCalls);
    }

    [Fact]
    public async Task ApplyOcrAsync_HoldsNonConcurrentEngineGateUntilTimedOutCallSettles() {
        var engine = new NonCooperativeSerialOcrEngine();
        var timeoutOptions = new OfficeDocumentOcrExecutionOptions {
            CandidateTimeout = TimeSpan.FromMilliseconds(20),
            ContinueOnError = true
        };

        try {
            OfficeDocumentOcrExecutionResult first = await CreateDocument(1).ApplyOcrAsync(engine, timeoutOptions);
            Assert.Contains(first.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout");

            Task<OfficeDocumentOcrExecutionResult> second = CreateDocument(1).ApplyOcrAsync(
                engine,
                new OfficeDocumentOcrExecutionOptions { CandidateTimeout = TimeSpan.FromSeconds(2) });
            Task earlyStart = await Task.WhenAny(engine.SecondCallStarted, Task.Delay(TimeSpan.FromMilliseconds(100)));
            Assert.NotSame(engine.SecondCallStarted, earlyStart);

            engine.CompleteFirstCall();
            Task completed = await Task.WhenAny(second, Task.Delay(TimeSpan.FromSeconds(2)));
            Assert.Same(second, completed);
            await second;

            Assert.Equal(1, engine.MaximumConcurrentCalls);
        } finally {
            engine.CompleteFirstCall();
        }
    }

    [Fact]
    public async Task ApplyOcrAsync_DoesNotStartAnotherCandidateWhileTimedOutSerialCallRuns() {
        var engine = new NonCooperativeSerialOcrEngine();

        try {
            OfficeDocumentOcrExecutionResult execution = await CreateDocument(2).ApplyOcrAsync(
                engine,
                new OfficeDocumentOcrExecutionOptions {
                    CandidateTimeout = TimeSpan.FromMilliseconds(20),
                    ContinueOnError = true
                });

            Assert.Equal(1, engine.CallCount);
            Assert.Equal(1, execution.Report.AttemptedCandidateCount);
            Assert.Equal(1, execution.Report.FailedCandidateCount);
            Assert.Equal(1, execution.Report.SkippedCandidateCount);
            Assert.Equal(1, execution.Report.InputBytes);
            Assert.Equal(2, execution.Diagnostics.Count(diagnostic => diagnostic.Code == "ocr-engine-timeout"));
            Assert.Equal(1, engine.MaximumConcurrentCalls);
        } finally {
            engine.CompleteFirstCall();
        }
    }

    [Fact]
    public async Task ApplyOcrAsync_DoesNotExceedParallelismWhenConcurrentEngineIgnoresTimeout() {
        var engine = new NonCooperativeConcurrentOcrEngine();

        try {
            OfficeDocumentOcrExecutionResult execution = await CreateDocument(4).ApplyOcrAsync(
                engine,
                new OfficeDocumentOcrExecutionOptions {
                    CandidateTimeout = TimeSpan.FromMilliseconds(20),
                    ContinueOnError = true,
                    MaxDegreeOfParallelism = 2
                });

            Assert.Equal(2, engine.CallCount);
            Assert.Equal(2, engine.MaximumConcurrentCalls);
            Assert.Equal(2, execution.Report.AttemptedCandidateCount);
            Assert.Equal(2, execution.Report.FailedCandidateCount);
            Assert.Equal(2, execution.Report.SkippedCandidateCount);
        } finally {
            engine.CompleteCalls();
        }
    }

    [Fact]
    public async Task ApplyOcrAsync_DoesNotStartQueuedCandidatesAfterSerialFailFastFailure() {
        var engine = new FailFastConcurrentOcrEngine();
        Task<OfficeDocumentOcrExecutionResult> execution = CreateDocument(5).ApplyOcrAsync(
            engine,
            new OfficeDocumentOcrExecutionOptions {
                ContinueOnError = false,
                MaxDegreeOfParallelism = 1
            });

        await engine.FirstCallStarted;
        engine.FailFirstCall();

        await Assert.ThrowsAsync<InvalidOperationException>(() => execution);
        Assert.Equal(1, engine.CallCount);
    }

    [Fact]
    public async Task ApplyOcrAsync_WaitsForStartedCandidatesAfterParallelFailFastFailure() {
        var engine = new FailFastConcurrentOcrEngine();
        Task<OfficeDocumentOcrExecutionResult> execution = CreateDocument(3).ApplyOcrAsync(
            engine,
            new OfficeDocumentOcrExecutionOptions {
                ContinueOnError = false,
                MaxDegreeOfParallelism = 2
            });

        try {
            await engine.TwoCallsStarted;
            engine.FailFirstCall();
            await Task.Delay(50);

            Assert.False(execution.IsCompleted);
            Assert.Equal(2, engine.CallCount);

            engine.CompleteRemainingCalls();
            await Assert.ThrowsAsync<InvalidOperationException>(() => execution);
            Assert.Equal(TaskStatus.RanToCompletion, engine.RemainingCallsCompleted.Status);
        } finally {
            engine.CompleteRemainingCalls();
        }
    }

    [Fact]
    public async Task OfficeDocumentOcrProcessor_FreezesOptionsForAsyncReaderPipeline() {
        var options = new OfficeDocumentOcrExecutionOptions { MaxCandidates = 1 };
        var processor = new OfficeDocumentOcrProcessor(new RecordingOcrEngine(), options);
        options.MaxCandidates = 2;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddProcessor(processor).Build();

        OfficeDocumentProcessingResult processing = await reader.ProcessDocumentAsync(CreateDocument(2));

        Assert.True(processing.Succeeded);
        Assert.Equal(1, processing.Document.Blocks.Count(block => block.Kind == "ocr-text"));
        Assert.Single(processing.Document.OcrCandidates);
        Assert.Equal("1", Assert.Single(processing.Document.Metadata, item => item.Id == "reader-ocr-execution-attempted-count").Value);
    }

    private static OfficeDocumentReadResult CreateDocument(int count) {
        var assets = new List<OfficeDocumentAsset>();
        var candidates = new List<OfficeDocumentOcrCandidate>();
        var diagnostics = new List<OfficeDocumentDiagnostic>();
        var pages = new List<OfficeDocumentPage>();
        for (int index = 1; index <= count; index++) {
            byte[] payload = new[] { (byte)index };
            string assetId = "asset-" + index;
            var location = new ReaderLocation { Path = "scan.pdf", Page = index, SourceBlockKind = "image", BlockAnchor = assetId };
            assets.Add(new OfficeDocumentAsset {
                Id = assetId,
                Kind = "image",
                MediaType = "image/png",
                Extension = ".png",
                LengthBytes = payload.LongLength,
                PayloadBytes = payload,
                PayloadHash = OfficeDocumentAssetHash.ComputeSha256Hex(payload),
                Location = location
            });
            var candidate = new OfficeDocumentOcrCandidate {
                Id = "ocr-" + index,
                Kind = "image",
                AssetId = assetId,
                Location = location,
                Region = new OfficeDocumentRegion { X = 0, Y = 0, Width = 10, Height = 10 }
            };
            candidates.Add(candidate);
            diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Ocr,
                Code = "ocr-needed",
                Message = "OCR needed.",
                Location = location
            });
            pages.Add(new OfficeDocumentPage { Number = index, Location = new ReaderLocation { Path = "scan.pdf", Page = index }, OcrCandidates = new[] { candidate } });
        }
        return new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Source = new OfficeDocumentSource { Path = "scan.pdf", SourceId = "scan" },
            Assets = assets,
            OcrCandidates = candidates,
            Diagnostics = diagnostics,
            Pages = pages
        };
    }

    private sealed class NonCooperativeSerialOcrEngine : IOfficeOcrEngine {
        private readonly TaskCompletionSource<OfficeOcrEngineResult> _firstCall = new TaskCompletionSource<OfficeOcrEngineResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _secondCallStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _activeCalls;
        private int _callCount;
        private int _maximumConcurrentCalls;

        public string Id => "non-cooperative-serial-fixture";

        public OfficeOcrEngineCapabilities Capabilities { get; } = new OfficeOcrEngineCapabilities {
            SupportedMediaTypes = new[] { "image/*" },
            SupportsConcurrentRequests = false
        };

        internal int MaximumConcurrentCalls => _maximumConcurrentCalls;

        internal int CallCount => _callCount;

        internal Task SecondCallStarted => _secondCallStarted.Task;

        internal void CompleteFirstCall() {
            _firstCall.TrySetResult(new OfficeOcrEngineResult { Text = "first" });
        }

        public async ValueTask<OfficeOcrEngineResult> RecognizeAsync(OfficeOcrEngineRequest request, CancellationToken cancellationToken = default) {
            int call = Interlocked.Increment(ref _callCount);
            int active = Interlocked.Increment(ref _activeCalls);
            while (true) {
                int current = _maximumConcurrentCalls;
                if (active <= current || Interlocked.CompareExchange(ref _maximumConcurrentCalls, active, current) == current) break;
            }
            try {
                if (call == 1) return await _firstCall.Task.ConfigureAwait(false);
                _secondCallStarted.TrySetResult(null);
                return new OfficeOcrEngineResult { Text = "second" };
            } finally {
                Interlocked.Decrement(ref _activeCalls);
            }
        }
    }

    private sealed class NonCooperativeConcurrentOcrEngine : IOfficeOcrEngine {
        private readonly TaskCompletionSource<OfficeOcrEngineResult> _completion = new TaskCompletionSource<OfficeOcrEngineResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _activeCalls;
        private int _callCount;
        private int _maximumConcurrentCalls;

        public string Id => "non-cooperative-concurrent-fixture";

        public OfficeOcrEngineCapabilities Capabilities { get; } = new OfficeOcrEngineCapabilities {
            SupportedMediaTypes = new[] { "image/*" },
            SupportsConcurrentRequests = true
        };

        internal int CallCount => _callCount;

        internal int MaximumConcurrentCalls => _maximumConcurrentCalls;

        internal void CompleteCalls() {
            _completion.TrySetResult(new OfficeOcrEngineResult { Text = "late" });
        }

        public async ValueTask<OfficeOcrEngineResult> RecognizeAsync(OfficeOcrEngineRequest request, CancellationToken cancellationToken = default) {
            Interlocked.Increment(ref _callCount);
            int active = Interlocked.Increment(ref _activeCalls);
            while (true) {
                int current = _maximumConcurrentCalls;
                if (active <= current || Interlocked.CompareExchange(ref _maximumConcurrentCalls, active, current) == current) break;
            }
            try {
                return await _completion.Task.ConfigureAwait(false);
            } finally {
                Interlocked.Decrement(ref _activeCalls);
            }
        }
    }

    private sealed class FailFastConcurrentOcrEngine : IOfficeOcrEngine {
        private readonly TaskCompletionSource<object?> _failFirstCall = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _firstCallStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _remainingCallsCompleted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _releaseRemainingCalls = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _twoCallsStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _callCount;

        public string Id => "fail-fast-concurrent-fixture";

        public OfficeOcrEngineCapabilities Capabilities { get; } = new OfficeOcrEngineCapabilities {
            SupportedMediaTypes = new[] { "image/*" },
            SupportsConcurrentRequests = true
        };

        internal int CallCount => _callCount;

        internal Task FirstCallStarted => _firstCallStarted.Task;

        internal Task RemainingCallsCompleted => _remainingCallsCompleted.Task;

        internal Task TwoCallsStarted => _twoCallsStarted.Task;

        internal void CompleteRemainingCalls() {
            _releaseRemainingCalls.TrySetResult(null);
        }

        internal void FailFirstCall() {
            _failFirstCall.TrySetResult(null);
        }

        public async ValueTask<OfficeOcrEngineResult> RecognizeAsync(OfficeOcrEngineRequest request, CancellationToken cancellationToken = default) {
            int call = Interlocked.Increment(ref _callCount);
            if (call == 1) {
                _firstCallStarted.TrySetResult(null);
                await _failFirstCall.Task.ConfigureAwait(false);
                throw new InvalidOperationException("Provider failure.");
            }
            _twoCallsStarted.TrySetResult(null);
            try {
                await _releaseRemainingCalls.Task.ConfigureAwait(false);
                return new OfficeOcrEngineResult { Text = "recognized" };
            } finally {
                _remainingCallsCompleted.TrySetResult(null);
            }
        }
    }

    private sealed class RecordingOcrEngine : IOfficeOcrEngine {
        private readonly List<string> _candidateIds = new List<string>();
        private int _activeCalls;
        private int _maximumConcurrentCalls;

        internal RecordingOcrEngine(bool supportsConcurrentRequests = true) {
            Capabilities = new OfficeOcrEngineCapabilities {
                SupportedMediaTypes = new[] { "image/*" },
                SupportsLineSpans = true,
                SupportsWordSpans = true,
                SupportsCharacterSpans = true,
                SupportsConfidence = true,
                SupportsConcurrentRequests = supportsConcurrentRequests
            };
        }

        public string Id => "fixture-engine";

        public OfficeOcrEngineCapabilities Capabilities { get; }

        internal IReadOnlyList<string> CandidateIds {
            get { lock (_candidateIds) return _candidateIds.ToArray(); }
        }

        internal int MaximumConcurrentCalls => _maximumConcurrentCalls;

        public async ValueTask<OfficeOcrEngineResult> RecognizeAsync(OfficeOcrEngineRequest request, CancellationToken cancellationToken = default) {
            lock (_candidateIds) _candidateIds.Add(request.Candidate.Id);
            int active = Interlocked.Increment(ref _activeCalls);
            while (true) {
                int current = _maximumConcurrentCalls;
                if (active <= current || Interlocked.CompareExchange(ref _maximumConcurrentCalls, active, current) == current) break;
            }
            try {
                await Task.Delay(request.Candidate.Id == "ocr-1" ? 40 : 5, cancellationToken);
                string text = "Text for " + request.Candidate.Id;
                return new OfficeOcrEngineResult {
                    Text = text,
                    Confidence = 0.9,
                    Language = request.Language,
                    Spans = new[] {
                        new OfficeOcrTextSpan { Sequence = 0, Level = OfficeOcrTextSpanLevel.Line, Text = text },
                        new OfficeOcrTextSpan { Sequence = 1, Level = OfficeOcrTextSpanLevel.Word, Text = "Text" },
                        new OfficeOcrTextSpan { Sequence = 2, Level = OfficeOcrTextSpanLevel.Character, Text = "T" }
                    }
                };
            } finally {
                Interlocked.Decrement(ref _activeCalls);
            }
        }
    }
}

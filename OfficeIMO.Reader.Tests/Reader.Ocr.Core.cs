using OfficeIMO.Ocr;
using OfficeIMO.Reader;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrCoreTests {
    [Fact]
    public void OcrContracts_AreNeutralAndReaderExecutionStaysInTheOptionalIntegration() {
        Assert.Empty(typeof(IOcrEngine).Assembly.GetReferencedAssemblies()
            .Where(reference => reference.Name?.StartsWith("OfficeIMO.", StringComparison.Ordinal) == true));
        Assert.Null(typeof(OfficeDocumentReadResult).Assembly.GetType("OfficeIMO.Reader.IOfficeOcrEngine"));
        Assert.Equal("OfficeIMO.Reader.Ocr", typeof(OfficeDocumentOcrExecutionExtensions).Assembly.GetName().Name);
    }

    [Fact]
    public void DelegateOcrEngine_RejectsOversizedRawIdentifierBeforeNormalization() {
        Assert.Throws<ArgumentException>(() => new DelegateOcrEngine(
            new string(' ', OcrEngineRunner.MaximumEngineIdCharacters) + "x",
            (_, _) => Task.FromResult(new OcrResult())));
    }

    [Fact]
    public void OcrProviderEntryGate_RejectsWorkAtTheDeadlineAndAfterCallerCancellation() {
        var expired = new OcrProviderEntryGate(
            System.Diagnostics.Stopwatch.StartNew(),
            TimeSpan.Zero,
            CancellationToken.None);
        Assert.False(expired.TryStart());
        Assert.False(expired.HasStarted);

        using var cancellation = new CancellationTokenSource();
        var canceled = new OcrProviderEntryGate(
            System.Diagnostics.Stopwatch.StartNew(),
            TimeSpan.FromMinutes(1),
            cancellation.Token);
        cancellation.Cancel();
        Assert.False(canceled.TryStart());
        Assert.False(canceled.HasStarted);

        var admitted = new OcrProviderEntryGate(
            System.Diagnostics.Stopwatch.StartNew(),
            TimeSpan.FromMinutes(1),
            CancellationToken.None);
        Assert.True(admitted.TryStart());
        admitted.SuppressIfNotStarted();
        Assert.True(admitted.HasStarted);
    }

    [Fact]
    public async Task ApplyOcrAsync_CapturesEngineIdentityAndCapabilitiesOncePerDocumentOperation() {
        OfficeDocumentReadResult source = CreateDocument(2);
        var engine = new ChangingIdentityOcrEngine();

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine);

        Assert.Equal(1, engine.IdReadCount);
        Assert.Equal(1, engine.CapabilitiesReadCount);
        Assert.Equal("snapshot-engine", execution.Report.EngineId);
        Assert.Equal(2, execution.Report.RecognizedCandidateCount);
    }

    [Fact]
    public async Task ApplyOcrAsync_PreservesCandidateOrderAndDetailedSpansUnderConcurrency() {
        OfficeDocumentReadResult source = CreateDocument(2);
        var engine = new RecordingOcrEngine(requiredConcurrentCalls: 2);

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
        Assert.Contains(execution.Recognitions[0].Result.Spans, span => span.Level == OcrTextSpanLevel.Line);
        Assert.Contains(execution.Recognitions[0].Result.Spans, span => span.Level == OcrTextSpanLevel.Word);
        Assert.Contains(execution.Recognitions[0].Result.Spans, span => span.Level == OcrTextSpanLevel.Character);
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
        string oversizedHierarchyId = new string('x', 257);
        var engine = new DelegateOcrEngine("bounded-fixture", (request, cancellationToken) => Task.FromResult(new OcrResult {
            Text = "1234567890",
            Confidence = 1.5,
            Spans = new[] {
                new OcrTextSpan {
                    Sequence = 0,
                    Level = OcrTextSpanLevel.Line,
                    Text = "1234567890",
                    Confidence = -0.5,
                    BlockId = oversizedHierarchyId,
                    ParagraphId = oversizedHierarchyId,
                    LineId = oversizedHierarchyId
                },
                new OcrTextSpan { Sequence = 1, Level = OcrTextSpanLevel.Word, Text = "12345" },
                new OcrTextSpan { Sequence = 2, Level = OcrTextSpanLevel.Character, Text = "1" }
            }
        }));

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxRecognizedCharactersPerCandidate = 5,
            MaxSpansPerCandidate = 2
        });

        Assert.Equal("12345", Assert.Single(execution.Document.Blocks, block => block.Kind == "ocr-text").Text);
        OcrResult result = Assert.Single(execution.Recognitions).Result;
        Assert.Equal(1D, result.Confidence);
        Assert.Equal(0D, result.Spans[0].Confidence);
        Assert.Null(result.Spans[0].BlockId);
        Assert.Null(result.Spans[0].ParagraphId);
        Assert.Null(result.Spans[0].LineId);
        Assert.Equal(2, result.Spans.Count);
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-text-limit");
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-span-limit");
        Assert.Single(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-confidence-out-of-range");
        Assert.Single(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-hierarchy-id-limit");
    }

    [Fact]
    public async Task ApplyOcrAsync_BoundsAllRetainedProviderControlledTextAndDiagnostics() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var engine = new DelegateOcrEngine("bounded-output-fixture", (_, _) => Task.FromResult(new OcrResult {
            Text = "recognized",
            Provider = "provider-name",
            Model = "provider-model",
            Language = "provider-language",
            Spans = new[] {
                new OcrTextSpan { Sequence = 0, Level = OcrTextSpanLevel.Word, Text = "abcdef" },
                new OcrTextSpan { Sequence = 1, Level = OcrTextSpanLevel.Word, Text = "ghijkl" }
            },
            Diagnostics = new[] {
                new OcrDiagnostic {
                    Code = "warning",
                    Message = "provider message",
                    Source = "provider source",
                    Attributes = new Dictionary<string, string> {
                        ["key"] = "value",
                        ["second"] = "attribute"
                    }
                },
                new OcrDiagnostic { Code = "second", Message = "discarded" }
            }
        }));

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxSpanCharactersPerCandidate = 6,
            MaxResultMetadataCharactersPerCandidate = 5,
            MaxProviderDiagnosticsPerCandidate = 1,
            MaxProviderDiagnosticCharactersPerCandidate = 8,
            MaxProviderDiagnosticAttributesPerCandidate = 1,
            MaxProviderDiagnosticAttributeCharactersPerCandidate = 4
        });

        OcrResult result = Assert.Single(execution.Recognitions).Result;
        Assert.Equal("provi", result.Provider);
        Assert.Null(result.Model);
        Assert.Null(result.Language);
        Assert.Equal("abcdef", result.Spans[0].Text);
        Assert.Equal(string.Empty, result.Spans[1].Text);
        OcrDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.True((diagnostic.Code.Length + diagnostic.Message.Length + (diagnostic.Source?.Length ?? 0)) <= 8);
        KeyValuePair<string, string> attribute = Assert.Single(diagnostic.Attributes);
        Assert.True(attribute.Key.Length + attribute.Value.Length <= 4);
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-result-metadata-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-span-text-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-text-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-text-limit");
    }

    [Fact]
    public async Task ApplyOcrAsync_BoundsRawProviderStringsBeforeTrimmingThem() {
        OfficeDocumentReadResult source = CreateDocument(1);
        string padded = new string(' ', 1024) + "unbounded-tail";
        var engine = new DelegateOcrEngine("raw-bounds-fixture", (_, _) => Task.FromResult(new OcrResult {
            Text = padded,
            Provider = padded,
            Spans = new[] {
                new OcrTextSpan {
                    Sequence = 0,
                    Level = OcrTextSpanLevel.Word,
                    Text = padded,
                    BlockId = padded
                }
            },
            Diagnostics = new[] {
                new OcrDiagnostic {
                    Code = padded,
                    Message = padded,
                    Source = padded,
                    Attributes = new Dictionary<string, string> { [padded] = padded }
                }
            }
        }));

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxRecognizedCharactersPerCandidate = 16,
            MaxSpanCharactersPerCandidate = 16,
            MaxResultMetadataCharactersPerCandidate = 16,
            MaxProviderDiagnosticCharactersPerCandidate = 16,
            MaxProviderDiagnosticAttributeCharactersPerCandidate = 16
        });

        OcrResult result = Assert.Single(execution.Recognitions).Result;
        Assert.Equal(string.Empty, result.Text);
        Assert.Null(result.Provider);
        Assert.Equal(string.Empty, Assert.Single(result.Spans).Text);
        Assert.Null(result.Spans[0].BlockId);
        OcrDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(string.Empty, diagnostic.Code);
        Assert.Equal(string.Empty, diagnostic.Message);
        Assert.Null(diagnostic.Source);
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-text-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-result-metadata-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-span-text-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-hierarchy-id-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-text-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-text-limit");
    }

    [Fact]
    public async Task ApplyOcrAsync_DoesNotReportTruncationWhenDiagnosticAttributeBudgetIsExactlyFilled() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var engine = new DelegateOcrEngine("exact-attribute-budget-fixture", (_, _) => Task.FromResult(new OcrResult {
            Text = "recognized",
            Diagnostics = new[] {
                new OcrDiagnostic {
                    Code = "notice",
                    Message = "message",
                    Attributes = new Dictionary<string, string> { ["key"] = "v" }
                }
            }
        }));

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxProviderDiagnosticAttributesPerCandidate = 1,
            MaxProviderDiagnosticAttributeCharactersPerCandidate = 4
        });

        KeyValuePair<string, string> attribute = Assert.Single(Assert.Single(execution.Recognitions).Result.Diagnostics[0].Attributes);
        Assert.Equal("key", attribute.Key);
        Assert.Equal("v", attribute.Value);
        Assert.DoesNotContain(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-limit");
        Assert.DoesNotContain(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-text-limit");
    }

    [Fact]
    public async Task ApplyOcrAsync_ConvertsPerCandidateTimeoutToRecoverableDiagnostic() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var engine = new DelegateOcrEngine("slow-fixture", async (request, cancellationToken) => {
            await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
            return new OcrResult { Text = "late" };
        });

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            CandidateTimeout = TimeSpan.FromMilliseconds(20),
            ContinueOnError = true
        });

        Assert.Equal(1, execution.Report.FailedCandidateCount + execution.Report.SkippedCandidateCount);
        Assert.Equal(0, execution.Report.RecognizedCandidateCount);
        Assert.Empty(execution.Recognitions);
        Assert.Single(execution.Document.OcrCandidates);
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout" && diagnostic.IsRecoverable == true);
    }

    [Fact]
    public async Task ApplyOcrAsync_ArmsTimeoutBeforeInvokingSynchronousProviderWork() {
        OfficeDocumentReadResult source = CreateDocument(1);
        using var providerInvoked = new ManualResetEventSlim(false);
        using var releaseProvider = new ManualResetEventSlim(false);
        using var cancellationObserved = new ManualResetEventSlim(false);
        var providerFinished = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var engine = new DelegateOcrEngine("synchronous-fixture", (_, cancellationToken) => {
            providerInvoked.Set();
            try {
                Assert.True(releaseProvider.Wait(TimeSpan.FromSeconds(10)));
                if (cancellationToken.WaitHandle.WaitOne(TimeSpan.Zero)) {
                    cancellationObserved.Set();
                }
                cancellationToken.ThrowIfCancellationRequested();
                return Task.FromResult(new OcrResult { Text = "late" });
            } finally {
                providerFinished.TrySetResult(null);
            }
        });

        try {
            Task<OfficeDocumentOcrExecutionResult> executionTask = source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
                CandidateTimeout = TimeSpan.FromMilliseconds(20),
                ContinueOnError = true,
            });

            Assert.True(providerInvoked.Wait(TimeSpan.FromSeconds(10)));
            OfficeDocumentOcrExecutionResult execution = await executionTask;
            releaseProvider.Set();
            Task completed = await Task.WhenAny(providerFinished.Task, Task.Delay(TimeSpan.FromSeconds(10)));

            Assert.Same(providerFinished.Task, completed);
            Assert.True(cancellationObserved.IsSet);
            Assert.Equal(1, execution.Report.FailedCandidateCount);
            Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout");
        } finally {
            releaseProvider.Set();
        }
    }

    [Fact]
    public async Task ApplyOcrAsync_PreservesTimeoutsBeyondTheSignedWaitBoundary() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var providerStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var cancellation = new CancellationTokenSource();
        var engine = new DelegateOcrEngine("long-timeout-fixture", async (_, cancellationToken) => {
            providerStarted.TrySetResult(null);
            await Task.Delay(Timeout.Infinite, cancellationToken);
            return new OcrResult { Text = "late" };
        });

        Task<OfficeDocumentOcrExecutionResult> execution = source.ApplyOcrAsync(
            engine,
            new OfficeDocumentOcrExecutionOptions { CandidateTimeout = TimeSpan.FromDays(30) },
            cancellation.Token);
        Task started = await Task.WhenAny(providerStarted.Task, Task.Delay(TimeSpan.FromSeconds(30)));
        Assert.Same(providerStarted.Task, started);

        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => execution);
    }

    [Fact]
    public async Task ApplyOcrAsync_EnforcesTimeoutWhenSynchronousEngineIgnoresCancellation() {
        OfficeDocumentReadResult source = CreateDocument(1);
        using var releaseProvider = new ManualResetEventSlim(false);
        var engine = new DelegateOcrEngine("synchronous-non-cooperative-fixture", (_, _) => {
            releaseProvider.Wait();
            return Task.FromResult(new OcrResult { Text = "late" });
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
            Assert.Equal(1, execution.Report.FailedCandidateCount + execution.Report.SkippedCandidateCount);
            Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout");
        } finally {
            releaseProvider.Set();
        }
    }

    [Fact]
    public async Task ApplyOcrAsync_DoesNotWaitForBlockingProviderCancellationCallback() {
        OfficeDocumentReadResult source = CreateDocument(1);
        using var callbackEntered = new ManualResetEventSlim(false);
        using var releaseCallback = new ManualResetEventSlim(false);
        var providerStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var providerCompletion = new TaskCompletionSource<OcrResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        var engine = new DelegateOcrEngine("blocking-cancellation-fixture", (request, cancellationToken) => {
            _ = cancellationToken.Register(() => {
                callbackEntered.Set();
                releaseCallback.Wait();
            });
            providerStarted.TrySetResult(null);
            return providerCompletion.Task;
        });

        try {
            Task<OfficeDocumentOcrExecutionResult> executionTask = source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
                CandidateTimeout = TimeSpan.FromMilliseconds(100),
                ContinueOnError = true
            });
            Task started = await Task.WhenAny(providerStarted.Task, Task.Delay(TimeSpan.FromSeconds(10)));
            Assert.Same(providerStarted.Task, started);

            Task completed = await Task.WhenAny(executionTask, Task.Delay(TimeSpan.FromSeconds(2)));
            Assert.Same(executionTask, completed);
            OfficeDocumentOcrExecutionResult execution = await executionTask;
            Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout");
            Assert.True(callbackEntered.Wait(TimeSpan.FromSeconds(2)));
        } finally {
            releaseCallback.Set();
            providerCompletion.TrySetResult(new OcrResult { Text = "late" });
        }
    }

    [Fact]
    public async Task ApplyOcrAsync_HoldsNonConcurrentGateUntilProviderCancellationCallbackSettles() {
        OfficeDocumentReadResult source = CreateDocument(1);
        using var callbackEntered = new ManualResetEventSlim(false);
        using var releaseCallback = new ManualResetEventSlim(false);
        var firstProviderStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var secondProviderStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var firstProviderCompletion = new TaskCompletionSource<OcrResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        int callCount = 0;
        var engine = new DelegateOcrEngine("cancellation-gate-fixture", (request, cancellationToken) => {
            int call = Interlocked.Increment(ref callCount);
            if (call == 1) {
                _ = cancellationToken.Register(() => {
                    firstProviderCompletion.TrySetResult(new OcrResult { Text = "late" });
                    callbackEntered.Set();
                    releaseCallback.Wait();
                });
                firstProviderStarted.TrySetResult(null);
                return firstProviderCompletion.Task;
            }

            secondProviderStarted.TrySetResult(null);
            return Task.FromResult(new OcrResult { Text = "second" });
        });

        try {
            OfficeDocumentOcrExecutionResult first = await source.ApplyOcrAsync(
                engine,
                new OfficeDocumentOcrExecutionOptions {
                    CandidateTimeout = TimeSpan.FromMilliseconds(100),
                    ContinueOnError = true
                });
            Assert.Same(
                firstProviderStarted.Task,
                await Task.WhenAny(firstProviderStarted.Task, Task.Delay(TimeSpan.FromSeconds(10))));
            Assert.Contains(first.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout");
            Assert.True(callbackEntered.Wait(TimeSpan.FromSeconds(2)));

            Task<OfficeDocumentOcrExecutionResult> second = source.ApplyOcrAsync(
                engine,
                new OfficeDocumentOcrExecutionOptions { CandidateTimeout = TimeSpan.FromSeconds(5) });
            Assert.NotSame(
                secondProviderStarted.Task,
                await Task.WhenAny(secondProviderStarted.Task, Task.Delay(TimeSpan.FromMilliseconds(200))));

            releaseCallback.Set();
            Assert.Same(second, await Task.WhenAny(second, Task.Delay(TimeSpan.FromSeconds(10))));
            OfficeDocumentOcrExecutionResult secondResult = await second;
            Assert.Equal(1, secondResult.Report.RecognizedCandidateCount);
            Assert.Equal(2, Volatile.Read(ref callCount));
        } finally {
            releaseCallback.Set();
            firstProviderCompletion.TrySetResult(new OcrResult { Text = "late" });
        }
    }

    [Fact]
    public async Task ApplyOcrAsync_EnforcesTimeoutWhenEngineIgnoresCancellation() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var completion = new TaskCompletionSource<OcrResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        var engine = new DelegateOcrEngine(
            "non-cooperative-fixture",
            (_, _) => completion.Task);

        try {
            Task<OfficeDocumentOcrExecutionResult> executionTask = source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
                CandidateTimeout = TimeSpan.FromMilliseconds(20),
                ContinueOnError = true
            });
            Task completed = await Task.WhenAny(executionTask, Task.Delay(TimeSpan.FromSeconds(2)));

            Assert.Same(executionTask, completed);
            OfficeDocumentOcrExecutionResult execution = await executionTask;
            Assert.Equal(1, execution.Report.FailedCandidateCount + execution.Report.SkippedCandidateCount);
            Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout");
        } finally {
            completion.TrySetResult(new OcrResult { Text = "late" });
        }
    }

    [Fact]
    public async Task ApplyOcrAsync_RemovesNonFiniteConfidenceAndNullProviderDiagnostics() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var engine = new DelegateOcrEngine("permissive-fixture", (_, _) => Task.FromResult(new OcrResult {
            Text = "recognized",
            Confidence = double.NaN,
            Spans = new[] {
                new OcrTextSpan { Sequence = 0, Level = OcrTextSpanLevel.Word, Text = "recognized", Confidence = double.PositiveInfinity }
            },
            Diagnostics = new OcrDiagnostic[] {
                null!,
                new OcrDiagnostic { Severity = OcrDiagnosticSeverity.Warning, Code = "provider-warning", Message = "Provider warning." }
            }
        }));

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine);

        OcrResult result = Assert.Single(execution.Recognitions).Result;
        Assert.Null(result.Confidence);
        Assert.Null(Assert.Single(result.Spans).Confidence);
        OcrDiagnostic providerDiagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("provider-warning", providerDiagnostic.Code);
        OfficeDocumentDiagnostic mappedDiagnostic = Assert.Single(execution.Diagnostics, diagnostic => diagnostic.Code == "provider-warning");
        Assert.Equal(OfficeDocumentDiagnosticCategory.Ocr, mappedDiagnostic.Category);
        Assert.Equal("permissive-fixture", mappedDiagnostic.Source);
        Assert.NotNull(mappedDiagnostic.Location);
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
            CandidateTimeout = TimeSpan.FromSeconds(2),
            ContinueOnError = true
        };

        try {
            Task<OfficeDocumentOcrExecutionResult> firstExecution =
                CreateDocument(1).ApplyOcrAsync(engine, timeoutOptions);
            Task firstCallStarted = await Task.WhenAny(
                engine.FirstCallStarted,
                Task.Delay(TimeSpan.FromSeconds(10)));
            Assert.Same(engine.FirstCallStarted, firstCallStarted);
            OfficeDocumentOcrExecutionResult first = await firstExecution;
            Assert.Contains(first.Diagnostics, diagnostic => diagnostic.Code == "ocr-engine-timeout");

            Task<OfficeDocumentOcrExecutionResult> second = CreateDocument(1).ApplyOcrAsync(
                engine,
                new OfficeDocumentOcrExecutionOptions { CandidateTimeout = TimeSpan.FromSeconds(2) });
            Task earlyStart = await Task.WhenAny(engine.SecondCallStarted, Task.Delay(TimeSpan.FromMilliseconds(100)));
            Assert.NotSame(engine.SecondCallStarted, earlyStart);

            engine.CompleteFirstCall();
            Task completed = await Task.WhenAny(second, Task.Delay(TimeSpan.FromSeconds(10)));
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
            Task<OfficeDocumentOcrExecutionResult> executionTask = CreateDocument(2).ApplyOcrAsync(
                engine,
                new OfficeDocumentOcrExecutionOptions {
                    CandidateTimeout = TimeSpan.FromSeconds(2),
                    ContinueOnError = true
                });
            Task firstCallStarted = await Task.WhenAny(
                engine.FirstCallStarted,
                Task.Delay(TimeSpan.FromSeconds(10)));
            Assert.Same(engine.FirstCallStarted, firstCallStarted);
            OfficeDocumentOcrExecutionResult execution = await executionTask;

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
            Task<OfficeDocumentOcrExecutionResult> executionTask = CreateDocument(4).ApplyOcrAsync(
                engine,
                new OfficeDocumentOcrExecutionOptions {
                    CandidateTimeout = TimeSpan.FromSeconds(5),
                    ContinueOnError = true,
                    MaxDegreeOfParallelism = 2
                });
            Task twoCallsStarted = await Task.WhenAny(
                engine.TwoCallsStarted,
                Task.Delay(TimeSpan.FromSeconds(10)));
            Assert.Same(engine.TwoCallsStarted, twoCallsStarted);
            OfficeDocumentOcrExecutionResult execution = await executionTask;

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

    private sealed class ChangingIdentityOcrEngine : IOcrEngine {
        private int _idReadCount;
        private int _capabilitiesReadCount;

        internal int IdReadCount => Volatile.Read(ref _idReadCount);
        internal int CapabilitiesReadCount => Volatile.Read(ref _capabilitiesReadCount);

        public string Id => Interlocked.Increment(ref _idReadCount) == 1
            ? "snapshot-engine"
            : new string('x', OcrEngineRunner.MaximumEngineIdCharacters + 1);

        public OcrEngineCapabilities Capabilities {
            get {
                Interlocked.Increment(ref _capabilitiesReadCount);
                return new OcrEngineCapabilities { SupportsConcurrentRequests = true };
            }
        }

        public Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default) =>
            Task.FromResult(new OcrResult { Text = "text" });
    }

    private sealed class NonCooperativeSerialOcrEngine : IOcrEngine {
        private readonly TaskCompletionSource<OcrResult> _firstCall = new TaskCompletionSource<OcrResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _firstCallStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _secondCallStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _activeCalls;
        private int _callCount;
        private int _maximumConcurrentCalls;

        public string Id => "non-cooperative-serial-fixture";

        public OcrEngineCapabilities Capabilities { get; } = new OcrEngineCapabilities {
            SupportedMediaTypes = new[] { "image/*" },
            SupportsConcurrentRequests = false
        };

        internal int MaximumConcurrentCalls => _maximumConcurrentCalls;

        internal int CallCount => _callCount;

        internal Task FirstCallStarted => _firstCallStarted.Task;

        internal Task SecondCallStarted => _secondCallStarted.Task;

        internal void CompleteFirstCall() {
            _firstCall.TrySetResult(new OcrResult { Text = "first" });
        }

        public async Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default) {
            int call = Interlocked.Increment(ref _callCount);
            int active = Interlocked.Increment(ref _activeCalls);
            while (true) {
                int current = _maximumConcurrentCalls;
                if (active <= current || Interlocked.CompareExchange(ref _maximumConcurrentCalls, active, current) == current) break;
            }
            try {
                if (call == 1) {
                    _firstCallStarted.TrySetResult(null);
                    return await _firstCall.Task.ConfigureAwait(false);
                }
                _secondCallStarted.TrySetResult(null);
                return new OcrResult { Text = "second" };
            } finally {
                Interlocked.Decrement(ref _activeCalls);
            }
        }
    }

    private sealed class NonCooperativeConcurrentOcrEngine : IOcrEngine {
        private readonly TaskCompletionSource<OcrResult> _completion = new TaskCompletionSource<OcrResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _twoCallsStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _activeCalls;
        private int _callCount;
        private int _maximumConcurrentCalls;

        public string Id => "non-cooperative-concurrent-fixture";

        public OcrEngineCapabilities Capabilities { get; } = new OcrEngineCapabilities {
            SupportedMediaTypes = new[] { "image/*" },
            SupportsConcurrentRequests = true
        };

        internal int CallCount => _callCount;

        internal int MaximumConcurrentCalls => _maximumConcurrentCalls;

        internal Task TwoCallsStarted => _twoCallsStarted.Task;

        internal void CompleteCalls() {
            _completion.TrySetResult(new OcrResult { Text = "late" });
        }

        public async Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default) {
            int callCount = Interlocked.Increment(ref _callCount);
            if (callCount >= 2) {
                _twoCallsStarted.TrySetResult(null);
            }
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

    private sealed class FailFastConcurrentOcrEngine : IOcrEngine {
        private readonly TaskCompletionSource<object?> _failFirstCall = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _firstCallStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _remainingCallsCompleted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _releaseRemainingCalls = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _twoCallsStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _callCount;

        public string Id => "fail-fast-concurrent-fixture";

        public OcrEngineCapabilities Capabilities { get; } = new OcrEngineCapabilities {
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

        public async Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default) {
            int call = Interlocked.Increment(ref _callCount);
            if (call == 1) {
                _firstCallStarted.TrySetResult(null);
                await _failFirstCall.Task.ConfigureAwait(false);
                throw new InvalidOperationException("Provider failure.");
            }
            _twoCallsStarted.TrySetResult(null);
            try {
                await _releaseRemainingCalls.Task.ConfigureAwait(false);
                return new OcrResult { Text = "recognized" };
            } finally {
                _remainingCallsCompleted.TrySetResult(null);
            }
        }
    }

    private sealed class RecordingOcrEngine : IOcrEngine {
        private readonly List<string> _candidateIds = new List<string>();
        private readonly int _requiredConcurrentCalls;
        private readonly TaskCompletionSource<object?> _requiredConcurrentCallsStarted =
            new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _activeCalls;
        private int _maximumConcurrentCalls;
        private int _startedCalls;

        internal RecordingOcrEngine(bool supportsConcurrentRequests = true, int requiredConcurrentCalls = 0) {
            _requiredConcurrentCalls = requiredConcurrentCalls;
            Capabilities = new OcrEngineCapabilities {
                SupportedMediaTypes = new[] { "image/*" },
                SupportsLineSpans = true,
                SupportsWordSpans = true,
                SupportsCharacterSpans = true,
                SupportsConfidence = true,
                SupportsConcurrentRequests = supportsConcurrentRequests
            };
        }

        public string Id => "fixture-engine";

        public OcrEngineCapabilities Capabilities { get; }

        internal IReadOnlyList<string> CandidateIds {
            get { lock (_candidateIds) return _candidateIds.ToArray(); }
        }

        internal int MaximumConcurrentCalls => _maximumConcurrentCalls;

        public async Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default) {
            lock (_candidateIds) _candidateIds.Add(request.CandidateId!);
            int active = Interlocked.Increment(ref _activeCalls);
            while (true) {
                int current = _maximumConcurrentCalls;
                if (active <= current || Interlocked.CompareExchange(ref _maximumConcurrentCalls, active, current) == current) break;
            }
            try {
                if (_requiredConcurrentCalls > 0) {
                    int started = Interlocked.Increment(ref _startedCalls);
                    if (started >= _requiredConcurrentCalls) {
                        _requiredConcurrentCallsStarted.TrySetResult(null);
                    }

                    await Task.WhenAny(
                        _requiredConcurrentCallsStarted.Task,
                        Task.Delay(TimeSpan.FromSeconds(2), cancellationToken));
                }
                await Task.Delay(request.CandidateId! == "ocr-1" ? 40 : 5, cancellationToken);
                string text = "Text for " + request.CandidateId!;
                return new OcrResult {
                    Text = text,
                    Confidence = 0.9,
                    Language = request.Language,
                    Spans = new[] {
                        new OcrTextSpan { Sequence = 0, Level = OcrTextSpanLevel.Line, Text = text },
                        new OcrTextSpan { Sequence = 1, Level = OcrTextSpanLevel.Word, Text = "Text" },
                        new OcrTextSpan { Sequence = 2, Level = OcrTextSpanLevel.Character, Text = "T" }
                    }
                };
            } finally {
                Interlocked.Decrement(ref _activeCalls);
            }
        }
    }
}

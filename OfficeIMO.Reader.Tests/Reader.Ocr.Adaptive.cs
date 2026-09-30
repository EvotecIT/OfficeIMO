using OfficeIMO.Ocr;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class AdaptiveOcrTests {
    [Fact]
    public void WordEvidenceIgnoresOverallScoresAndCountsUnknownConfidenceAsUncertain() {
        var result = Words("A B C", .8, null, double.NaN);
        result.Confidence = 1;
        var quality = new OcrReviewPolicy(.8, 0).Assess(result);
        Assert.Equal(3, quality.WordCount);
        Assert.Equal(0, quality.LowConfidenceWordCount);
        Assert.Equal(2, quality.UnknownConfidenceWordCount);
        Assert.False(quality.MeetsThresholds);
        Assert.False(new OcrReviewPolicy().Assess(new OcrResult { Text = "A", Confidence = 1 }).MeetsThresholds);
    }

    [Fact]
    public async Task PassingBaselineDoesNotRunRetriesOrApproveCorrectness() {
        var engine = new AdaptiveOcrEngine("adaptive", new[] {
            Attempt("baseline", () => Words("A B", .9, .9)),
            Attempt("retry", () => throw new Exception("Should not execute"))
        });
        AdaptiveOcrResult result = await engine.RecognizeWithReviewAsync(Request());
        Assert.Single(result.Attempts);
        Assert.False(result.ReviewRecommended);
        Assert.Equal("adaptive-ocr-thresholds-met", result.Result.Diagnostics[0].Code);
        Assert.Contains("does not establish", result.Result.Diagnostics[0].Message);
    }

    [Fact]
    public async Task ImprovedUncertaintyCanSelectRetryButDisagreementStillRequiresReview() {
        var engine = new AdaptiveOcrEngine("adaptive", new[] {
            Attempt("baseline", () => Words("Total 10", .9, .1)),
            Attempt("retry", () => Words("Total 70", .9, .9))
        });
        AdaptiveOcrResult result = await engine.RecognizeWithReviewAsync(Request());
        Assert.Equal("Total 70", result.Result.Text);
        Assert.Equal(1, result.SelectedAttempt);
        Assert.True(result.Quality.MeetsThresholds);
        Assert.True(result.HasDisagreement);
        Assert.True(result.ReviewRecommended);
        Assert.Equal("adaptive-ocr-review-recommended", result.Result.Diagnostics[0].Code);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" \t")]
    public async Task PassingRetryRepairsMissingAggregateTextWithoutApprovingDisagreement(string missing) {
        var baseline = Words("A", 1); baseline.Text = missing;
        var engine = new AdaptiveOcrEngine("adaptive", new[] {
            Attempt("baseline", () => baseline), Attempt("retry", () => Words("A", 1))
        });
        AdaptiveOcrResult result = await engine.RecognizeWithReviewAsync(Request());
        Assert.Equal("A", result.Result.Text);
        Assert.Equal(1, result.SelectedAttempt);
        Assert.True(result.Quality.MeetsThresholds);
        Assert.True(result.HasDisagreement);
        Assert.True(result.ReviewRecommended);
    }

    [Fact]
    public async Task ShorterHighConfidenceRetryCannotDiscardBaselineCoverage() {
        var engine = new AdaptiveOcrEngine("adaptive", new[] {
            Attempt("baseline", () => Words("A B C D", .1, .1, .1, .1)),
            Attempt("retry", () => Words("A", 1))
        });
        AdaptiveOcrResult result = await engine.RecognizeWithReviewAsync(Request());
        Assert.Equal(0, result.SelectedAttempt);
        Assert.Equal("A B C D", result.Result.Text);
        Assert.True(result.ReviewRecommended);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FailedOrOversizedRetryRetainsBaselineAndContentFreeEvidence(bool oversized) {
        var engine = new AdaptiveOcrEngine("adaptive", new[] {
            Attempt("baseline", () => Words("A B", .1, .1)),
            Attempt("retry", () => oversized ? Words(new string('x', 100), 1) : throw new Exception("PRIVATE_DOCUMENT_SENTINEL"))
        }, maximumTextCharacters: 20);
        AdaptiveOcrResult result = await engine.RecognizeWithReviewAsync(Request());
        Assert.Equal("A B", result.Result.Text);
        Assert.Equal("failed", result.Attempts[1].Outcome);
        Assert.True(result.RetryIncomplete);
        Assert.DoesNotContain("PRIVATE_DOCUMENT_SENTINEL", System.Text.Json.JsonSerializer.Serialize(result.Attempts));
    }

    [Fact]
    public async Task AttemptsReceiveIndependentInputsAndSelectedGeometryRemainsOwned() {
        var source = Request(); source.Payload = new byte[] { 1, 2 }; source.Region = new OcrRegion { Width = 10 };
        var span = new OcrTextSpan { Level = OcrTextSpanLevel.Word, Text = "A", Confidence = 1, Region = new OcrRegion { X = 5, Width = 3 } };
        var baseline = new DelegateOcrEngine("baseline", (request, _) => {
            request.Payload[0] = 99; request.Region!.Width = 999;
            return Task.FromResult(Words("A", .1));
        });
        var retry = new DelegateOcrEngine("retry", (request, _) => {
            Assert.Equal(1, request.Payload[0]); Assert.Equal(10, request.Region!.Width);
            return Task.FromResult(new OcrResult { Text = "A", Spans = new[] { span } });
        });
        var engine = new AdaptiveOcrEngine("adaptive", new[] { new OcrRecognitionAttempt("baseline", baseline), new OcrRecognitionAttempt("retry", retry) });
        AdaptiveOcrResult result = await engine.RecognizeWithReviewAsync(source);
        span.Region!.X = 999;
        Assert.Equal(5, result.Result.Spans[0].Region!.X);
        Assert.Equal(1, source.Payload[0]); Assert.Equal(10, source.Region.Width);
        Assert.False(result.HasDisagreement);
    }

    [Fact]
    public async Task CallerCancellationDuringRetryNeverReturnsBaselineAsSuccess() {
        using var cancellation = new CancellationTokenSource();
        var entered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var retry = new DelegateOcrEngine("retry", async (_, token) => {
            entered.TrySetResult(true);
            await Task.Delay(Timeout.Infinite, token);
            return Words("A", 1);
        });
        var engine = new AdaptiveOcrEngine("adaptive", new[] { Attempt("baseline", () => Words("A", .1)), new OcrRecognitionAttempt("retry", retry) });
        Task<AdaptiveOcrResult> task = engine.RecognizeWithReviewAsync(Request(), cancellation.Token);
        Assert.Same(entered.Task, await Task.WhenAny(entered.Task, Task.Delay(TimeSpan.FromSeconds(10))));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => task);
    }

    [Fact]
    public async Task WarningAndOmittedEvidenceCannotPassThresholdsThroughSharedRunner() {
        var engine = new AdaptiveOcrEngine("adaptive", new[] { Attempt("baseline", () => Words("A B", 1, 1)) }, maximumSpans: 1);
        OcrResult result = await OcrEngineRunner.RecognizeAsync(engine, Request(), TimeSpan.FromSeconds(10));
        Assert.Equal("adaptive-ocr-review-recommended", result.Diagnostics[0].Code);
        Assert.Equal("false", result.Diagnostics[0].Attributes["thresholds-met"]);
        Assert.Equal(1, result.OmittedSpanCount);
        Assert.False(new OcrReviewPolicy().Assess(result).MeetsThresholds);
    }

    [Fact]
    public async Task RetryTimeoutRetainsReviewableBaselineAndKeepsUnderlyingProviderGate() {
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var entered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var retry = new DelegateOcrEngine("retry", async (_, _) => {
            entered.TrySetResult(true); await release.Task; return Words("A", 1);
        });
        var engine = new AdaptiveOcrEngine("adaptive", new[] {
            Attempt("baseline", () => Words("A", .1)), new OcrRecognitionAttempt("retry", retry)
        }, timeout: TimeSpan.FromSeconds(1));
        Task<AdaptiveOcrResult> pending = engine.RecognizeWithReviewAsync(Request());
        try {
            Assert.Same(entered.Task, await Task.WhenAny(entered.Task, Task.Delay(TimeSpan.FromSeconds(10))));
            AdaptiveOcrResult result = await pending;
            Assert.Equal("A", result.Result.Text); Assert.True(result.RetryIncomplete);
            Assert.Equal("timed-out", result.Attempts[1].Outcome);
            OcrEngineTimeoutException error = await Assert.ThrowsAsync<OcrEngineTimeoutException>(() =>
                OcrEngineRunner.RecognizeAsync(retry, Request(), TimeSpan.FromMilliseconds(50)));
            Assert.False(error.ProviderCallStarted);
        } finally { release.TrySetResult(true); }
    }

    [Fact]
    public void ReviewPolicyRejectsInvalidThresholdsAndWarningsRemainReviewEvidence() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OcrReviewPolicy(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OcrReviewPolicy(maximumUncertainWordFraction: 2));
        var evidence = Words("A", 1);
        evidence.Diagnostics = new[] { new OcrDiagnostic { Severity = OcrDiagnosticSeverity.Warning } };
        Assert.False(new OcrReviewPolicy().Assess(evidence).MeetsThresholds);
    }

    [Fact]
    public async Task NestedRunnerPreservesPreviouslyOmittedSpansDiagnosticsAndAttributes() {
        OcrResult evidence = Words("A B", 1, 1);
        evidence.Diagnostics = new[] {
            new OcrDiagnostic { Attributes = new System.Collections.Generic.Dictionary<string, string> { ["a"] = "1", ["b"] = "2" } },
            new OcrDiagnostic()
        };
        var provider = new DelegateOcrEngine("provider", (_, _) => Task.FromResult(evidence));
        var nested = new DelegateOcrEngine("nested", (request, token) => OcrEngineRunner.CreateExecution(provider)
            .RecognizeAsync(request, TimeSpan.FromSeconds(10), new OcrResultCaptureLimits(1, 1, 1), token));
        OcrResult result = await OcrEngineRunner.RecognizeAsync(nested, Request(), TimeSpan.FromSeconds(10));
        Assert.Equal(1, result.OmittedSpanCount);
        Assert.Equal(1, result.OmittedDiagnosticCount);
        Assert.Equal(1, Assert.Single(result.Diagnostics).OmittedAttributeCount);
    }

    [Fact]
    public async Task InvalidUnicodeRetryRetainsBaselineAndRequiresReview() {
        var engine = new AdaptiveOcrEngine("adaptive", new[] {
            Attempt("baseline", () => Words("A", .1)), Attempt("retry", () => Words("\uD800", 1))
        });
        AdaptiveOcrResult result = await engine.RecognizeWithReviewAsync(Request());
        Assert.Equal("A", result.Result.Text);
        Assert.Equal("failed", result.Attempts[1].Outcome);
        Assert.True(result.RetryIncomplete);
    }

    [Fact]
    public async Task OrientationSharesInputBudgetAndReceivesOwnedPayload() {
        bool called = false;
        var provider = new DelegateOcrEngine("provider", (request, _) => {
            called = true; request.Payload[0] = 99;
            return Task.FromResult(new OcrResult());
        }, new OcrEngineCapabilities { SupportsOrientationDetection = true });
        var engine = new AdaptiveOcrEngine("adaptive", new[] { new OcrRecognitionAttempt("baseline", provider) }, maximumInputBytes: 1);
        var request = Request(); request.Operation = OcrOperation.DetectOrientation; request.Payload = new byte[] { 1, 2 };
        await Assert.ThrowsAsync<ArgumentException>(() => engine.RecognizeAsync(request));
        Assert.False(called);
        request.Payload = new byte[] { 1 };
        await engine.RecognizeAsync(request);
        Assert.True(called); Assert.Equal(1, request.Payload[0]);
    }

    private static OcrRequest Request() => new() { Payload = new byte[] { 1 }, MediaType = "image/png" };
    private static OcrRecognitionAttempt Attempt(string name, Func<OcrResult> run) => new(name, new DelegateOcrEngine(name, (_, _) => Task.FromResult(run())));
    private static OcrResult Words(string text, params double?[] confidence) => new() {
        Text = text, Spans = text.Split(' ').Select((word, index) => new OcrTextSpan {
            Text = word, Level = OcrTextSpanLevel.Word, Confidence = index < confidence.Length ? confidence[index] : null
        }).ToArray()
    };
}

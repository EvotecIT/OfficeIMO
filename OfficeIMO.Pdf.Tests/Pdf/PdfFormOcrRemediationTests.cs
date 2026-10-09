using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFormOcrRemediationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task FormPreparationStaysAttachedDuringRecognitionAndOrientationCleanup(bool concurrent, bool orientation) {
        var source = PdfDocument.Load(PdfDocument.Create().TextField("Name", width: 180, height: 24).ToBytes());
        var original = source.ToBytes();
        var entered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var cleaning = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var cancellation = new CancellationTokenSource();
        OcrOperation? actual = null;
        var engine = new DelegateOcrEngine("form-lifetime", async (request, token) => {
            actual = request.Operation; entered.TrySetResult(true);
            try { await Task.Delay(Timeout.Infinite, token); }
            finally { cleaning.TrySetResult(true); await release.Task; }
            return new OcrResult();
        }, new OcrEngineCapabilities { SupportsConcurrentRequests = concurrent, SupportsOrientationDetection = orientation });
        var preparation = source.PrepareFormOcrAsync(engine, new() { Dpi = 72, DetectOrientation = orientation }, cancellation.Token);
        try {
            Assert.Same(entered.Task, await Task.WhenAny(entered.Task, Task.Delay(TimeSpan.FromSeconds(10)))); await entered.Task; cancellation.Cancel();
            Assert.Same(cleaning.Task, await Task.WhenAny(cleaning.Task, Task.Delay(TimeSpan.FromSeconds(10)))); await cleaning.Task; await Task.Delay(100);
            Assert.False(preparation.IsCompleted);
            Assert.Equal(orientation ? OcrOperation.DetectOrientation : OcrOperation.RecognizeText, actual);
            release.TrySetResult(true); await Assert.ThrowsAnyAsync<OperationCanceledException>(() => preparation);
            Assert.Equal(original, source.ToBytes());
        } finally {
            release.TrySetResult(true); cancellation.Cancel();
            try { await preparation; } catch (OperationCanceledException) { }
        }
    }

    [Theory]
    [InlineData("1,1,2", 3, new[] { 1, 2 })]
    [InlineData("2,1-2", 2, new[] { 2, 1 })]
    public async Task RepeatedPhysicalPagesAreRecognizedOnceWithoutChangingCallerSelection(string selection, int maxPages, int[] expected) {
        var source = PdfFormOcrReviewTests.Fixture("reportlab-scanned-form.pdf");
        var original = source.ToBytes();
        var selected = PdfPageSelection.Parse(selection);
        var options = new PdfOcrMergeOptions {
            ReadOptions = new() { PageSelection = selected }, Dpi = 72,
            MaxPages = maxPages, MaxConcurrentPages = 1
        };
        var calls = new List<int>();
        var geometry = PdfFormOcrReviewTests.Engine(source);
        var engine = new DelegateOcrEngine("repeated-pages", (request, token) => {
            calls.Add(request.PageNumber!.Value);
            return geometry.RecognizeAsync(request, token);
        });
        var review = await source.PrepareFormOcrAsync(engine, options);
        Assert.Equal(expected, calls);
        Assert.Equal(2, review.Ocr.Pages.Count);
        Assert.Equal(7, review.Proposals.Count);
        Assert.All(review.Proposals, proposal => Assert.False(proposal.IsAmbiguous));
        Assert.Same(selected, options.ReadOptions.PageSelection);
        Assert.Equal(original, source.ToBytes());
    }

    [Fact]
    public async Task AppendOnlyReviewedFillRegeneratesOriginalProducerAppearanceWithoutRewritingSignedPrefix() {
        var source = PdfFormOcrReviewTests.Fixture("reportlab-certified-form.pdf");
        var original = source.ToBytes();
        var field = source.Inspect().FormFieldsByName["FullName"];
        var widget = Assert.Single(field.Widgets);
        var bounds = source.GetPageLayouts()[0].MapUserSpaceRectangleToVisual(widget.X1, widget.Y1, widget.X2, widget.Y2);
        var engine = new DelegateOcrEngine("certified-appearance", (request, _) => Task.FromResult(new OcrResult {
            Spans = request.PageNumber == 1 ? [new OcrTextSpan {
                Text = "Alex Morgan", Confidence = .99, Level = OcrTextSpanLevel.Word,
                CoordinateUnit = OcrCoordinateUnit.Points,
                Region = new() { X = bounds.Left + 2, Y = bounds.Top + 2, Width = bounds.Width - 4, Height = bounds.Height - 4 }
            }] : []
        }));
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, source.PlanMutation(PdfMutationOperation.FillFormFields, ["FullName"]).ExecutionMode);
        var review = await source.PrepareFormOcrAsync(engine, new() { Dpi = 72 });
        var result = review.Apply(source, new Dictionary<PdfFormOcrProposal, PdfFormFieldValue> {
            [Assert.Single(review.Proposals)] = "Reviewed name"
        });
        var updated = result.ToBytes();
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_FORM_OCR_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output);
            File.WriteAllBytes(Path.Combine(output, "certified-source.pdf"), original);
            File.WriteAllBytes(Path.Combine(output, "certified-review.pdf"), updated);
        }
        Assert.True(updated.AsSpan(0, original.Length).SequenceEqual(original));
        Assert.Equal("Reviewed name", result.Inspect().FormFieldsByName["FullName"].Value);
        Assert.Equal(false, result.Inspect().AcroFormNeedAppearances);
        Assert.Contains("<5265766965776564206E616D65> Tj", PdfEncoding.Latin1GetString(updated, original.Length, updated.Length - original.Length));
        Assert.Equal(original, source.ToBytes());
    }
}

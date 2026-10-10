using System.Text;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed partial class PdfWorkspaceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CertifiedManualFormFillRegeneratesAppearanceForBatchAndSingleField(bool batch) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-certified-form-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "PdfFormOcr", "reportlab-certified-form.pdf"), source);
        byte[] original = File.ReadAllBytes(source);
        try {
            using var workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            if (batch) await workspace.FillFormFieldsAsync(new Dictionary<string, PdfFormFieldValue> { ["FullName"] = "Reviewed name" }, CancellationToken.None);
            else await workspace.FillFormFieldAsync("FullName", "Reviewed name", false, CancellationToken.None);
            string saved = Path.Combine(root, "filled-copy.pdf");
            await workspace.SaveAsync(saved, CancellationToken.None);
            byte[] updated = File.ReadAllBytes(saved);
            Assert.Equal(workspace.CopyBytes(), updated);
            string? output = Environment.GetEnvironmentVariable("OFFICEIMO_FORM_OCR_OUTPUT");
            if (!string.IsNullOrWhiteSpace(output)) {
                Directory.CreateDirectory(output);
                File.WriteAllBytes(Path.Combine(output, batch ? "certified-batch.pdf" : "certified-single.pdf"), updated);
            }
            Assert.Equal("Reviewed name", workspace.CreateDocumentSnapshot().Inspect().FormFieldsByName["FullName"].Value);
            Assert.True(updated.AsSpan(0, original.Length).SequenceEqual(original));
            Assert.Equal(false, workspace.CreateDocumentSnapshot().Inspect().AcroFormNeedAppearances);
            Assert.Contains("<5265766965776564206E616D65> Tj", Encoding.Latin1.GetString(updated.AsSpan(original.Length)));
            Assert.Single(workspace.Journal);
            await workspace.UndoAsync(CancellationToken.None);
            Assert.Equal(original, workspace.CopyBytes());
            Assert.Equal(original, File.ReadAllBytes(source));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task FormRecognitionOwnsCpuAdmissionThroughProviderCancellationCleanupAndQueuedMutation() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-form-async-owner-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf"), other = Path.Combine(root, "other.pdf");
        var fixture = Path.Combine(AppContext.BaseDirectory, "Fixtures", "PdfFormOcr", "reportlab-scanned-form.pdf");
        File.Copy(fixture, source); File.Copy(fixture, other);
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var cleaning = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var cancellation = new CancellationTokenSource();
        using var mutationCancellation = new CancellationTokenSource();
        using var queuedCancellation = new CancellationTokenSource();
        Task<PdfWorkspaceFormOcrReview>? preparation = null;
        Task? mutation = null;
        Task<int>? queued = null;
        Task<int>? following = null;
        var engine = new DelegateOcrEngine("async-owner-fixture", async (_, token) => {
            entered.TrySetResult();
            try { await Task.Delay(Timeout.Infinite, token); }
            finally { cleaning.TrySetResult(); await release.Task; }
            return new OcrResult();
        });
        try {
            using var workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            using var otherWorkspace = await PdfWorkspace.OpenAsync(other, CancellationToken.None);
            byte[] original = workspace.CopyBytes();
            preparation = workspace.PrepareFormOcrAsync(engine, new() { Dpi = 72, MaxConcurrentPages = 1 }, cancellation.Token);
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(30));
            // This source supports form filling; its existing scripts deliberately block page rewrites.
            mutation = workspace.FillFormFieldAsync("FullName", "Pending correction", false, mutationCancellation.Token);
            int concurrentStarts = 0;
            queued = otherWorkspace.RunCancellableCpuWorkAsync(() => Interlocked.Increment(ref concurrentStarts), queuedCancellation.Token);
            cancellation.Cancel();
            await cleaning.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await Task.Delay(100);
            queuedCancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => queued);
            mutationCancellation.Cancel();
            try { await mutation; } catch (OperationCanceledException) { }
            Assert.False(preparation.IsCompleted);
            Assert.Equal(0, concurrentStarts);
            Assert.Equal(original, workspace.CopyBytes());
            Assert.Empty(workspace.Journal);
            workspace.Dispose();
            following = otherWorkspace.RunCancellableCpuWorkAsync(() => 1, CancellationToken.None);
            await Task.Delay(100);
            Assert.False(following.IsCompleted);
            release.TrySetResult();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => preparation);
            Assert.Equal(1, await following.WaitAsync(TimeSpan.FromSeconds(10)));
            Assert.Equal(original, File.ReadAllBytes(source));
        } finally {
            release.TrySetResult(); cancellation.Cancel(); mutationCancellation.Cancel(); queuedCancellation.Cancel();
            if (preparation is not null) try { await preparation; } catch (OperationCanceledException) { }
            if (mutation is not null) try { await mutation; } catch (OperationCanceledException) { }
            if (queued is not null) try { await queued; } catch (OperationCanceledException) { }
            if (following is not null) try { await following; } catch (ObjectDisposedException) { }
            Directory.Delete(root, true);
        }
    }
}

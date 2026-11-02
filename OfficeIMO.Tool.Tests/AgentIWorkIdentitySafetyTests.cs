using System.Runtime.InteropServices;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Agent;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed partial class AgentIWorkConversionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Pdf_comparison_rejects_native_source_ids_before_opening_replacement_fifos(bool comparisonSource) {
        if (OperatingSystem.IsWindows()) return;
        using var files = new Files();
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]));
        string nativeId = (await service.InspectAsync(files.Input)).SourceId;
        string pdf = files.Output("pdf");
        PdfDocument.Create(document => document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(150), global::OfficeIMO.OfficeLength.Points(180)))).Save(pdf);
        string pdfId = (await service.InspectAsync(pdf)).SourceId;
        File.Delete(files.Input);
        Assert.Equal(0, CreateIdentityFifo(files.Input, 0x180));

        var failure = await CompleteWithoutFifoWriter(files.Input, () =>
            service.ComparePdfAsync(comparisonSource ? pdfId : nativeId,
                comparisonSource ? nativeId : pdfId, files.Output("html"), new PdfWorkflowSettings()));
        Assert.Contains("Select one local PDF file", Assert.IsType<AgentUsageException>(failure).Message);
        Assert.False(File.Exists(files.Output("html")));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Pdf_comparison_reauthorizes_native_source_paths_before_identity_reads(bool directoryPackage,
        bool comparisonSource) {
        if (OperatingSystem.IsWindows()) return;
        using var files = new Files(directoryPackage);
        using var outside = new Files(directoryPackage);
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]));
        string nativeId = (await service.InspectAsync(files.Input)).SourceId;
        string pdf = files.Output("pdf");
        PdfDocument.Create(document => document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(150), global::OfficeIMO.OfficeLength.Points(180)))).Save(pdf);
        string pdfId = (await service.InspectAsync(pdf)).SourceId;
        if (directoryPackage) {
            Directory.Delete(files.Input, true);
            Directory.CreateSymbolicLink(files.Input, outside.Input);
        } else {
            File.Delete(files.Input);
            File.CreateSymbolicLink(files.Input, outside.Input);
        }

        await Assert.ThrowsAsync<UnauthorizedAccessException>(() => service.ComparePdfAsync(
            comparisonSource ? pdfId : nativeId, comparisonSource ? nativeId : pdfId,
            files.Output("html"), new PdfWorkflowSettings()));
        Assert.False(File.Exists(files.Output("html")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Native_registry_identity_reads_reject_replacement_fifos_and_honor_cancellation(bool cancelled) {
        if (OperatingSystem.IsWindows()) return;
        using var files = new Files();
        var registry = new AgentSourceRegistry();
        string sourceId = registry.Register(files.Input).SourceId;
        File.Delete(files.Input);
        Assert.Equal(0, CreateIdentityFifo(files.Input, 0x180));
        var failure = await CompleteWithoutFifoWriter(files.Input, () => {
            _ = registry.Resolve(sourceId, new CancellationToken(cancelled));
            return Task.CompletedTask;
        });
        if (cancelled) Assert.IsType<OperationCanceledException>(failure);
        else Assert.IsType<InvalidDataException>(failure);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Native_registry_rejects_relocated_source_paths_before_reading_the_new_target(bool directoryPackage) {
        if (OperatingSystem.IsWindows()) return;
        using var files = new Files(directoryPackage);
        using var outside = new Files(directoryPackage);
        var registry = new AgentSourceRegistry();
        string sourceId = registry.Register(files.Input).SourceId;
        string fifo;
        if (directoryPackage) {
            fifo = Path.Combine(outside.Input, "replacement.bin");
            Directory.Delete(files.Input, true);
            Directory.CreateSymbolicLink(files.Input, outside.Input);
        } else {
            fifo = outside.Input;
            File.Delete(outside.Input);
            File.Delete(files.Input);
            File.CreateSymbolicLink(files.Input, outside.Input);
        }
        Assert.Equal(0, CreateIdentityFifo(fifo, 0x180));
        var failure = await CompleteWithoutFifoWriter(fifo, () => {
            _ = registry.Resolve(sourceId);
            return Task.CompletedTask;
        });
        Assert.IsType<AgentUsageException>(failure);
    }

    private static async Task<Exception?> CompleteWithoutFifoWriter(string fifo, Func<Task> action) {
        Task<Exception?> attempt = Task.Run(() => Record.ExceptionAsync(action));
        bool completedWithoutWriter = ReferenceEquals(
            await Task.WhenAny(attempt, Task.Delay(TimeSpan.FromSeconds(2))), attempt);
        if (!completedWithoutWriter) {
            // Nonblocking read/write open releases a stuck reader without leaving
            // a second blocking writer behind, even if the attempt has just ended.
            int descriptor = OpenIdentityFifo(fifo, 2 | (OperatingSystem.IsMacOS() ? 4 : 2048));
            Assert.True(descriptor >= 0, "Could not release the replacement FIFO reader.");
            Assert.Equal(0, CloseIdentityFifo(descriptor));
        }
        Exception? failure = await attempt.WaitAsync(TimeSpan.FromSeconds(2));
        Assert.True(completedWithoutWriter, "Source identity resolution blocked waiting for a FIFO writer.");
        return failure;
    }

    [DllImport("libc", EntryPoint = "mkfifo", SetLastError = true)]
    private static extern int CreateIdentityFifo(string path, uint mode);

    [DllImport("libc", EntryPoint = "open", SetLastError = true)]
    private static extern int OpenIdentityFifo(string path, int flags);

    [DllImport("libc", EntryPoint = "close", SetLastError = true)]
    private static extern int CloseIdentityFifo(int descriptor);
}

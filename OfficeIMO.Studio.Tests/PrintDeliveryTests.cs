using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class PrintDeliveryTests {
    [Fact]
    public async Task TrayDiscoveryRecoveryClearsOnlyItsOwnError() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var printer = new RecordingPrinter { DelayPaperSources = true };
            using var model = Create(printer, new StudioJobHistory(StudioLocalization.Current));
            model.Status = "Reviewed sheets ready";
            model.SelectedPrinter = new("Unavailable", false, false);
            printer.Sources["Unavailable"].SetException(new IOException("Driver offline"));
            await model.PaperSourceDiscovery;
            Assert.True(model.HasPaperSourceError);
            Assert.Contains("Driver offline", model.PaperSourceError);
            Assert.Equal("Reviewed sheets ready", model.Status);
            model.SelectedPrinter = new("Available", false, false);
            printer.Sources["Available"].SetResult([new("tray-2", "Lower tray")]);
            await model.PaperSourceDiscovery;
            Assert.False(model.HasPaperSourceError);
            Assert.Equal("Reviewed sheets ready", model.Status);
            Assert.Contains(model.PaperSourceChoices, choice => choice.Id == "tray-2");
            return true;
        }, default);
    }

    [Fact]
    public async Task SwitchingPrintersDiscardsLateTrayDiscoveryAndClearsSelection() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var printer = new RecordingPrinter { DelayPaperSources = true };
            using var model = Create(printer, new StudioJobHistory(StudioLocalization.Current));
            model.SelectedPrinter = new("First", false, false);
            Task first = model.PaperSourceDiscovery;
            Assert.True(model.IsDiscoveringPaperSources);
            model.SelectedPrinter = new("Second", false, false);
            model.PrintOutputPath = "previous-driver.pdf";
            Task second = model.PaperSourceDiscovery;
            printer.Sources["Second"].SetResult([new("second-tray", "Second tray")]);
            await second;
            model.SelectedPaperSource = model.PaperSourceChoices.Single(choice => choice.Id == "second-tray");
            printer.Sources["First"].SetResult([new("first-tray", "First tray")]);
            await first;
            Assert.Equal("second-tray", model.SelectedPaperSource.Id);
            Assert.DoesNotContain(model.PaperSourceChoices, choice => choice.Id == "first-tray");
            Assert.False(model.IsDiscoveringPaperSources);
            model.SelectedPrinter = new("Third", false, false);
            Assert.Empty(model.PrintOutputPath);
            Task third = model.PaperSourceDiscovery;
            Assert.Null(model.SelectedPaperSource.Id);
            model.Dispose();
            printer.Sources["Third"].SetResult([new("third-tray", "Third tray")]);
            await third;
            Assert.DoesNotContain(model.PaperSourceChoices, choice => choice.Id == "third-tray");
            Assert.False(model.CanPrint);
            return true;
        }, default);
    }

    [Fact]
    public async Task DeliversReviewedSnapshotAndInvalidatesChangedSettings() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var printer = new RecordingPrinter();
            int reads = 0;
            var history = new StudioJobHistory(StudioLocalization.Current);
            using var model = Create(printer, history, () => reads++);
            await model.RefreshPrintersCommand.ExecuteAsync(null);
            model.Pages = "3,1";
            model.SelectedPagesPerSheet = model.PagesPerSheetChoices.Single(choice => choice.Value == 2);
            await model.BuildPreviewCommand.ExecuteAsync(null);
            Assert.True(model.CanPrint, model.Status);
            Assert.Single(model.Sheets);
            Assert.Equal(new[] { 3, 1 }, model.Sheets[0].Placements.Select(placement => placement.PageNumber));
            model.Copies = 2;
            model.SelectedPaperSource = model.PaperSourceChoices.Single(choice => choice.Id == "tray-2");
            model.SelectedDuplex = model.DuplexChoices.Single(choice => choice.Value == PdfPrintDuplex.LongEdge);
            await model.PrintCommand.ExecuteAsync(null);
            Assert.Equal(1, reads);
            Assert.NotNull(printer.Document);
            Assert.Equal(new[] { 3, 1 }, printer.Document.Plan.SelectedPages);
            Assert.Equal(2, printer.Options!.Copies);
            Assert.Equal("tray-2", printer.Options.PaperSourceId);
            Assert.Equal(PdfPrintDuplex.LongEdge, printer.Options.Duplex);
            Assert.Contains("Printer accepted job test-7", model.Status);
            Assert.False(history.Entries[0].HasOutput);
            Assert.False(history.Entries[0].IsActive);
            byte[] copy = printer.Document.Sheets[0].GetPng();
            byte original = copy[0];
            copy[0] ^= 255;
            Assert.Equal(original, printer.Document.Sheets[0].GetPng()[0]);
            model.PrintDpi = 300;
            Assert.False(model.HasPreview);
            Assert.False(model.CanPrint);
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancellationDistinguishesQueuedFromStartedDelivery(bool started) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var printer = new RecordingPrinter { WaitForCancellation = true };
            var history = new StudioJobHistory(StudioLocalization.Current);
            using var model = Create(printer, history);
            await model.RefreshPrintersCommand.ExecuteAsync(null);
            await model.BuildPreviewCommand.ExecuteAsync(null);
            using IDisposable? first = started ? null : await history.EnterAsync(CancellationToken.None);
            using IDisposable? second = started ? null : await history.EnterAsync(CancellationToken.None);
            Task printing = model.PrintCommand.ExecuteAsync(null);
            if (started) await printer.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.True(model.IsBusy);
            Assert.False(model.CanPrint);
            history.Entries[0].CancelCommand.Execute(null);
            await printing.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.False(model.IsBusy);
            Assert.False(history.Entries[0].IsActive);
            Assert.Equal(started ? "Check output" : "Cancelled", history.Entries[0].Status);
            Assert.Equal(started, printer.Document is not null);
            if (started) Assert.Contains("Check the printer queue", model.Status);
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(PdfPrintOrientation.Portrait)]
    [InlineData(PdfPrintOrientation.Landscape)]
    public async Task LargestPaperAtHighestOfferedResolutionPreparesAndDelivers(PdfPrintOrientation orientation) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var printer = new RecordingPrinter();
            using var model = Create(printer, new StudioJobHistory(StudioLocalization.Current));
            await model.RefreshPrintersCommand.ExecuteAsync(null);
            model.Pages = "1";
            model.SelectedPaper = model.PaperChoices.Single(choice => choice.Name == "A3");
            model.SelectedOrientation = model.OrientationChoices.Single(choice => choice.Value == orientation);
            model.PrintDpi = model.PrintDpiChoices.Max();
            await model.BuildPreviewCommand.ExecuteAsync(null);
            Assert.True(model.CanPrint, model.Status);
            await model.PrintCommand.ExecuteAsync(null);
            Assert.NotNull(printer.Document);
            Assert.Single(printer.Document.Sheets);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task AcceptedPrintWithCleanupWarningRemainsCompleted() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var printer = new RecordingPrinter { CleanupWarning = "Could not remove private print staging." };
            var history = new StudioJobHistory(StudioLocalization.Current);
            using var model = Create(printer, history);
            await model.RefreshPrintersCommand.ExecuteAsync(null);
            await model.BuildPreviewCommand.ExecuteAsync(null);
            await model.PrintCommand.ExecuteAsync(null);
            Assert.StartsWith("Printer accepted job test-7", model.Status);
            Assert.Contains(printer.CleanupWarning, model.Status);
            Assert.Equal("Completed", history.Entries[0].Status);
            return true;
        }, default);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PrinterDiscoveryPreservesAcceptedStatusAndPreventsConcurrentSubmission(bool empty) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var printer = new RecordingPrinter();
            using var model = Create(printer, new StudioJobHistory(StudioLocalization.Current));
            await model.RefreshPrintersCommand.ExecuteAsync(null);
            await model.BuildPreviewCommand.ExecuteAsync(null);
            await model.PrintCommand.ExecuteAsync(null);
            string accepted = model.Status;
            printer.PrinterDiscovery = new(TaskCreationOptions.RunContinuationsAsynchronously);
            Task refresh = model.RefreshPrintersCommand.ExecuteAsync(null);
            Assert.False(model.CanPrint);
            Assert.False(model.PrintCommand.CanExecute(null));
            if (empty) printer.PrinterDiscovery.SetResult([]);
            else printer.PrinterDiscovery.SetException(new IOException("Driver offline"));
            await refresh;
            Assert.True(model.HasPrinterDiscoveryError);
            Assert.Equal(accepted, model.Status);
            printer.PrinterDiscovery = null;
            await model.RefreshPrintersCommand.ExecuteAsync(null);
            Assert.False(model.HasPrinterDiscoveryError);
            Assert.True(model.CanPrint);
            Assert.Equal(accepted, model.Status);
            return true;
        }, default);
    }

    private static PrintPreviewViewModel Create(RecordingPrinter printer, StudioJobHistory history, Action? read = null) =>
        new(_ => Task.FromResult<string?>(null), null, readSnapshot: (_, _) => {
            read?.Invoke();
            return Task.FromResult(PdfDocument.Create(document => {
                for (int index = 1; index <= 3; index++) {
                    int number = index;
                    document.Page(page => page.Size(200, 300).Content(content =>
                        content.Item(item => item.Paragraph(paragraph => paragraph.Text("Reviewed page " + number)))));
                }
            }));
        }, printers: printer, jobHistory: history) { InputPath = Path.Combine(Path.GetTempPath(), "reviewed-snapshot.pdf") };

    private sealed class RecordingPrinter : IPdfPrinterService {
        public PdfPreparedPrintDocument? Document { get; private set; }
        public PdfPrintDeliveryOptions? Options { get; private set; }
        public bool WaitForCancellation { get; init; }
        public string? CleanupWarning { get; init; }
        public TaskCompletionSource<IReadOnlyList<PdfPrinterInfo>>? PrinterDiscovery { get; set; }
        public bool DelayPaperSources { get; init; }
        public Dictionary<string, TaskCompletionSource<IReadOnlyList<PdfPaperSourceInfo>>> Sources { get; } = new();
        public TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken cancellationToken = default) =>
            PrinterDiscovery?.Task ?? Task.FromResult<IReadOnlyList<PdfPrinterInfo>>([new("Test queue", true, false)]);
        public Task<IReadOnlyList<PdfPaperSourceInfo>> GetPaperSourcesAsync(string printerName, CancellationToken cancellationToken = default) {
            if (!DelayPaperSources) return Task.FromResult<IReadOnlyList<PdfPaperSourceInfo>>([new("tray-2", "Lower tray")]);
            var completion = new TaskCompletionSource<IReadOnlyList<PdfPaperSourceInfo>>(TaskCreationOptions.RunContinuationsAsynchronously);
            Sources.Add(printerName, completion);
            return completion.Task; // Deliberately model a driver that ignores cancellation.
        }
        public async Task<PdfPrintSubmission> SubmitAsync(PdfPreparedPrintDocument document, PdfPrintDeliveryOptions options, CancellationToken cancellationToken = default) {
            Document = document;
            Options = options;
            Started.TrySetResult();
            if (WaitForCancellation) {
                try { await Task.Delay(Timeout.Infinite, cancellationToken); }
                catch (OperationCanceledException error) { throw new PdfPrintDeliveryException("test-7", error); }
            }
            return new(options.PrinterName, "test-7", document.Sheets.Count, options.Copies, null) { CleanupWarning = CleanupWarning };
        }
    }
}

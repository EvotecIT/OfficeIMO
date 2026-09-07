using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioActiveCloseTests {
    [Theory]
    [InlineData("WaitAndClose", false)]
    [InlineData("CancelWorkAndClose", true)]
    [InlineData("KeepOpen", false)]
    public async Task WindowCloseOffersAnExplicitDecisionAndWaitsForActualQueuedWork(string choice, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            using var firstPermit = await services.Jobs.EnterAsync(CancellationToken.None);
            using var secondPermit = await services.Jobs.EnterAsync(CancellationToken.None);
            var window = new MainWindow(services) { Width = 960, Height = 620 };
            Task? running = null;
            try {
                window.Show();
                string source = Path.Combine(services.Paths.Root, "queued.html");
                File.WriteAllText(source, "<html><body><p>Wait before closing</p></body></html>");
                var conversion = window.ViewModel.ConversionWorkbench;
                conversion.OutputFolder = services.Paths.Root;
                conversion.Jobs.Add(new ConversionJobViewModel(source, conversion.Routes.Single(route => route.Route.Id == "html-pdf")));
                running = conversion.RunQueueCommand.ExecuteAsync(null);
                Assert.True(window.TabHost.HasBusyDocuments);
                window.Close();
                Assert.True(window.IsVisible);
                var dialog = Assert.Single(window.OwnedWindows.OfType<ActiveOperationsDialog>());
                Assert.True(conversion.IsBusy);
                Assert.Equal("Queued", services.Jobs.Entries[0].Status);
                LayoutAndCapture(dialog, choice + "-decision");
                Click(dialog, choice);
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                if (choice == "WaitAndClose") {
                    Assert.True(dialog.IsVisible);
                    Assert.True(window.IsVisible);
                    LayoutAndCapture(dialog, choice + "-waiting");
                } else if (choice == "KeepOpen") {
                    Assert.False(dialog.IsVisible);
                    Assert.True(conversion.IsBusy);
                }
                firstPermit.Dispose();
                secondPermit.Dispose();
                await running.WaitAsync(TimeSpan.FromSeconds(15));
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                Assert.False(conversion.IsBusy);
                Assert.Equal(choice == "KeepOpen", window.IsVisible);
                var job = Assert.Single(services.Jobs.Entries);
                Assert.False(job.IsActive);
                if (choice == "CancelWorkAndClose") {
                    Assert.Equal("Cancelled", job.Status);
                    Assert.False(File.Exists(Path.Combine(services.Paths.Root, "queued.pdf")));
                } else {
                    Assert.Equal("Completed", job.Status);
                    Assert.Equal(1, PdfDocument.Load(File.ReadAllBytes(job.OutputPath!)).Inspect().PageCount);
                }
            } finally {
                firstPermit.Dispose(); secondPermit.Dispose();
                window.TabHost.CancelAllOperations();
                if (running is not null) await running;
                foreach (Window dialog in window.OwnedWindows.ToArray()) dialog.Close(false);
                window.Close();
            }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ClosingOneBusyTabWaitsWithoutCancellingWorkInAnotherTab() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            using var firstPermit = await services.Jobs.EnterAsync(CancellationToken.None);
            using var secondPermit = await services.Jobs.EnterAsync(CancellationToken.None);
            var window = new MainWindow(services);
            Task? firstRun = null;
            Task? secondRun = null;
            try {
                window.Show();
                string source = Path.Combine(services.Paths.Root, "first.pdf");
                string other = Path.Combine(services.Paths.Root, "second.pdf");
                PdfDocument.Create(document => document.Page(page => page.Size(200, 300))).Save(source);
                File.Copy(source, other);
                await window.TabHost.OpenDocumentAsync(source);
                var firstTab = window.TabHost.SelectedTab!;
                await window.TabHost.OpenDocumentAsync(other);
                var secondTab = window.TabHost.SelectedTab!;
                firstTab.Document.DocumentHealth.InputPath = source;
                secondTab.Document.DocumentHealth.InputPath = other;
                firstRun = firstTab.Document.DocumentHealth.RunCommand.ExecuteAsync(null);
                secondRun = secondTab.Document.DocumentHealth.RunCommand.ExecuteAsync(null);
                Assert.True(firstTab.Document.CanCancelOperation);
                Assert.True(secondTab.Document.CanCancelOperation);
                Task closing = window.TabHost.CloseTabAsync(firstTab);
                var dialog = Assert.Single(window.OwnedWindows.OfType<ActiveOperationsDialog>());
                Click(dialog, "CancelWorkAndClose");
                await firstRun.WaitAsync(TimeSpan.FromSeconds(5));
                await closing.WaitAsync(TimeSpan.FromSeconds(5));
                Assert.Same(secondTab, Assert.Single(window.TabHost.Tabs));
                Assert.True(secondTab.Document.CanCancelOperation);
                Assert.True(window.IsVisible);
                firstPermit.Dispose(); secondPermit.Dispose();
                await secondRun.WaitAsync(TimeSpan.FromSeconds(10));
                Assert.Equal(1, services.Jobs.Entries.Count(entry => entry.Status == "Cancelled"));
                Assert.Equal(1, services.Jobs.Entries.Count(entry => entry.Status == "Completed"));
            } finally {
                firstPermit.Dispose(); secondPermit.Dispose();
                window.TabHost.CancelAllOperations();
                if (firstRun is not null) await firstRun;
                if (secondRun is not null) await secondRun;
                foreach (Window dialog in window.OwnedWindows.ToArray()) dialog.Close(false);
                window.Close();
            }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task InFlightCancellationKeepsCommittedOutputAndWaitsForPublicationCleanup() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var guard = new PauseSecondPublication();
            string root = Directory.CreateDirectory(Path.Combine(services.Paths.Root, "publication")).FullName;
            using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                services: services, publicationGuard: guard);
            var owner = new Window();
            Task? running = null;
            try {
                owner.Show();
                var conversion = document.ConversionWorkbench;
                conversion.OutputFolder = root;
                var route = conversion.Routes.Single(route => route.Route.Id == "html-pdf");
                foreach (string name in new[] { "committed", "cancelled" }) {
                    string source = Path.Combine(root, name + ".html");
                    File.WriteAllText(source, "<html><body><p>Safe close</p></body></html>");
                    conversion.Jobs.Add(new ConversionJobViewModel(source, route));
                }
                running = conversion.RunQueueCommand.ExecuteAsync(null);
                await guard.Reached.Task.WaitAsync(TimeSpan.FromSeconds(10));
                string committed = Path.Combine(root, "committed.pdf");
                byte[] bytes = File.ReadAllBytes(committed);
                var dialog = new ActiveOperationsDialog([document], services.Localizer);
                Task<bool> closing = dialog.ShowDialog<bool>(owner);
                Click(dialog, "CancelWorkAndClose");
                Assert.True(document.CanCancelOperation);
                Assert.False(closing.IsCompleted);
                guard.Release.TrySetResult();
                await running.WaitAsync(TimeSpan.FromSeconds(10));
                Assert.True(await closing.WaitAsync(TimeSpan.FromSeconds(5)));
                Assert.False(document.CanCancelOperation);
                Assert.Equal(bytes, File.ReadAllBytes(committed));
                Assert.Equal(1, PdfDocument.Load(bytes).Inspect().PageCount);
                Assert.False(File.Exists(Path.Combine(root, "cancelled.pdf")));
                Assert.Equal(new[] { "Completed", "Cancelled" }, conversion.Jobs.Select(job => job.Status));
                Assert.Equal(new[] { "cancelled.html", "committed.html", "committed.pdf" },
                    Directory.EnumerateFileSystemEntries(root).Select(Path.GetFileName).OrderBy(name => name, StringComparer.Ordinal));
            } finally {
                guard.Release.TrySetResult();
                document.CancelCurrentOperation();
                if (running is not null) await running;
                foreach (Window dialog in owner.OwnedWindows.ToArray()) dialog.Close(false);
                owner.Close();
            }
            return true;
        }, CancellationToken.None);
    }

    private sealed class PauseSecondPublication : IOfficeWorkflowPublicationGuard {
        private int _calls;
        internal TaskCompletionSource Reached { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) {
            if (Interlocked.Increment(ref _calls) == 2) {
                Reached.TrySetResult();
                // Simulate a guard completing asynchronous work before observing cancellation.
                await Release.Task;
            }
            return true;
        }
    }

    private static void Click(Window dialog, string name) =>
        dialog.GetVisualDescendants().OfType<Button>().Single(button => button.Name == name)
            .RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    private static void LayoutAndCapture(Window dialog, string name) {
        dialog.UpdateLayout();
        foreach (Button button in dialog.GetVisualDescendants().OfType<Button>()) {
            Assert.True(button.IsEffectivelyVisible);
            Point point = button.TranslatePoint(default, dialog)!.Value;
            Assert.InRange(point.X, 0, dialog.Bounds.Width - button.Bounds.Width);
            Assert.InRange(point.Y, 0, dialog.Bounds.Height - button.Bounds.Height);
        }
        using var frame = dialog.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, "active-close-" + name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}

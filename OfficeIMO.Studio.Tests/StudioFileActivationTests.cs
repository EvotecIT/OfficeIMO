using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioFileActivationTests {
    [Fact]
    public async Task Activation_waits_for_startup_and_serializes_PDF_tabs_and_Apple_intake_with_provider_access() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var numbers = new TestStorageFile("content://activation/numbers", File.ReadAllBytes(Fixture("simple.numbers")), "Opened.numbers");
            var first = new TestStorageFile("content://activation/first", StudioProviderDocumentTests.CreatePdf(), "First.pdf");
            var second = new TestStorageFile("content://activation/second", StudioProviderDocumentTests.CreatePdf(), "Second.pdf");
            var window = new MainWindow(services);
            try {
                var args = new FileActivatedEventArgs([numbers.Item, first.Item, first.Item]);
                Task initial = window.OpenActivatedItemsAsync(args.Files);
                Task subsequent = window.OpenActivatedItemsAsync([second.Item]);
                Assert.False(initial.IsCompleted);
                Assert.Empty(window.ViewModel.ConversionWorkbench.Jobs);
                window.Show();
                await Task.WhenAll(initial, subsequent);
                Assert.False(window.IsStartingUp);
                Assert.Equal(2, window.TabHost.Tabs.Count);
                Assert.Equal("Second.pdf", window.ViewModel.DocumentName);
                // Conversion work belongs to the tab where it was accepted; switching tabs must not lose it.
                var queue = window.TabHost.OperationDocuments.SelectMany(document => document.ConversionWorkbench.Jobs);
                var job = Assert.Single(queue);
                Assert.Equal("numbers-xlsx", job.Route.Route.Id);
                Assert.Equal(numbers.Bookmark, services.Storage.Describe(job.InputPath).Bookmark);
                Assert.Equal(0, numbers.Disposals);
                Assert.True(first.Reads > 0 && second.Reads > 0);
                Assert.Equal(first.Reads, first.ClosedReads);
                Assert.Equal(second.Reads, second.ClosedReads);
            } finally { window.Close(); }
            Assert.Equal(1, numbers.Disposals);
            Assert.Equal(1, first.Disposals);
            Assert.Equal(1, second.Disposals);
            return true;
        }, default);
    }

    [Fact]
    public async Task Closing_before_startup_releases_pending_activation_items_and_completes_batches() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var first = new TestStorageFile("content://activation/first", [], "First.numbers");
            var second = new TestStorageFile("content://activation/second", [], "Second.pages");
            var window = new MainWindow(services);
            Task initial = window.OpenActivatedItemsAsync([first.Item, first.Item]);
            Task next = window.OpenActivatedItemsAsync([second.Item]);
            window.Close();
            await Task.WhenAll(initial, next).WaitAsync(TimeSpan.FromSeconds(5));
            Assert.Equal(1, first.Disposals);
            Assert.Equal(1, second.Disposals);
            Assert.Equal(0, first.Reads + second.Reads);
            return true;
        }, default);
    }

    [Fact]
    public async Task Unsupported_activation_does_not_prevent_a_later_supported_batch() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var unsupported = new TestStorageFile("content://activation/unsupported", [], "Unknown.xyz");
            var valid = new TestStorageFile("content://activation/valid", File.ReadAllBytes(Fixture("simple.pages")), "Document.pages");
            var window = new MainWindow(services);
            try {
                window.Show();
                await window.OpenActivatedItemsAsync([unsupported.Item]);
                Assert.Equal(services.Localizer.Get("Activation.Unsupported"), window.ViewModel.ErrorMessage);
                Assert.Equal(1, unsupported.Disposals);
                Assert.Equal(0, unsupported.Reads);
                await window.OpenActivatedItemsAsync([valid.Item]);
                Assert.Equal("pages-docx", Assert.Single(window.ViewModel.ConversionWorkbench.Jobs).Route.Route.Id);
                Assert.Equal(0, valid.Disposals);
            } finally { window.Close(); }
            Assert.Equal(1, valid.Disposals);
            return true;
        }, default);
    }

    [Fact]
    public async Task Failed_permission_releases_the_item_and_continues_the_same_batch() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var denied = new TestStorageFile("content://activation/denied", [], "Denied.pages") { DenyBookmark = true };
            var valid = new TestStorageFile("content://activation/valid", File.ReadAllBytes(Fixture("simple.numbers")), "Document.numbers");
            var window = new MainWindow(services);
            try {
                window.Show();
                window.ViewModel.ShowConversionWorkbenchCommand.Execute(null);
                await window.OpenActivatedItemsAsync([denied.Item, valid.Item]);
                Assert.Equal(1, denied.Disposals);
                Assert.Equal("numbers-xlsx", Assert.Single(window.ViewModel.ConversionWorkbench.Jobs).Route.Route.Id);
                Assert.Equal(0, valid.Disposals);
                Assert.Equal("Provider permission expired.", window.ViewModel.ErrorMessage);
            } finally { window.Close(); }
            Assert.Equal(1, valid.Disposals);
            return true;
        }, default);
    }

    [Fact]
    public async Task Busy_activation_preserves_window_owned_access_and_releases_new_references() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var owned = new TestStorageFile("content://activation/owned", File.ReadAllBytes(Fixture("simple.pages")), "Document.pages");
            var rejected = new TestStorageFile("content://activation/rejected", [], "Later.numbers");
            var window = new MainWindow(services);
            try {
                window.Show();
                await window.OpenActivatedItemsAsync([owned.Item]);
                window.ViewModel.IsOpening = true;
                await window.OpenActivatedItemsAsync([owned.Item, rejected.Item]);
                Assert.Equal(services.Localizer.Get("Activation.Busy"), window.ViewModel.ErrorMessage);
                Assert.Equal(0, owned.Disposals);
                Assert.Equal(1, rejected.Disposals);
                Assert.Single(window.ViewModel.ConversionWorkbench.Jobs);
                window.ViewModel.IsOpening = false;
            } finally { window.Close(); }
            Assert.Equal(1, owned.Disposals);
            return true;
        }, default);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Successful_activation_clears_prior_intake_errors_but_preserves_other_errors(bool replaceWithOperationError) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var initial = new TestStorageFile("content://activation/initial", File.ReadAllBytes(Fixture("simple.pages")), "Initial.pages");
            var invalid = new TestStorageFile("content://activation/invalid", [], "Unknown.xyz");
            var valid = new TestStorageFile("content://activation/next", File.ReadAllBytes(Fixture("simple.numbers")), "Next.numbers");
            var window = new MainWindow(services);
            try {
                window.Show();
                await window.OpenActivatedItemsAsync([initial.Item]);
                Assert.True(window.ViewModel.IsConversionMode);
                await window.OpenActivatedItemsAsync([invalid.Item]);
                Assert.True(window.ViewModel.HasVisibleError);
                if (replaceWithOperationError) window.ViewModel.ErrorMessage = "Another operation failed.";
                await window.OpenActivatedItemsAsync([valid.Item]);
                Assert.Equal(2, window.ViewModel.ConversionWorkbench.Jobs.Count);
                if (replaceWithOperationError) Assert.Equal("Another operation failed.", window.ViewModel.ErrorMessage);
                else Assert.False(window.ViewModel.HasVisibleError);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    [Fact]
    public async Task All_Apple_startup_arguments_enter_the_queue() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var window = new MainWindow(services);
            try {
                window.OpenInitialDocument([Fixture("simple.pages"), Fixture("simple.numbers"), Fixture("tabledeck.key")]);
                window.Show();
                var deadline = DateTime.UtcNow.AddSeconds(5);
                while (window.IsStartingUp && DateTime.UtcNow < deadline) await Task.Delay(10);
                Assert.False(window.IsStartingUp);
                Assert.Equal(new[] { "pages-docx", "numbers-xlsx", "keynote-pptx" }, window.ViewModel.ConversionWorkbench.Jobs.Select(job => job.Route.Route.Id));
                Assert.Empty(window.TabHost.Tabs);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "IWork", name);
}

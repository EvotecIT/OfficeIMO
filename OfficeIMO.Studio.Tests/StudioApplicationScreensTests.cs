using Avalonia;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioApplicationScreensTests {
    [Fact]
    public async Task Selecting_a_tab_shows_its_document_and_closing_a_tab_keeps_the_application_screen() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var window = new MainWindow(services);
            try {
                window.Show();
                string first = CreateDocument(services.Paths.Root, "first.pdf");
                string second = CreateDocument(services.Paths.Root, "second.pdf");
                await window.TabHost.OpenDocumentAsync(first);
                var firstTab = window.TabHost.SelectedTab!;
                await window.TabHost.OpenDocumentAsync(second);
                Assert.True(window.ViewModel.IsPdfWorkspaceMode);

                window.ViewModel.ShowConversionWorkbenchCommand.Execute(null);
                Assert.True(window.ViewModel.IsConversionMode);

                // Choosing a document tab means "show me that document".
                window.TabHost.SelectedTab = firstTab;
                Assert.Same(firstTab.Document, window.ViewModel);
                Assert.True(window.ViewModel.IsPdfWorkspaceMode);

                // Closing a tab from an application screen leaves the user where they were.
                window.ViewModel.ShowJobsCommand.Execute(null);
                await CloseSelectedTabAsync(window);
                Assert.Single(window.TabHost.Tabs);
                Assert.True(window.ViewModel.IsJobsMode);

                // With no document left, a document view has nothing to show and Home takes over.
                window.ViewModel.ShowPdfWorkspaceCommand.Execute(null);
                await CloseSelectedTabAsync(window);
                Assert.Empty(window.TabHost.Tabs);
                Assert.True(window.ViewModel.IsHomeMode);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task Workbench_errors_appear_on_the_visible_document() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var window = new MainWindow(services);
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(CreateDocument(services.Paths.Root, "shown.pdf"));
                MainWindowViewModel owner = window.TabHost.OperationDocuments.Single(document => document.OwnsWorkbenches);
                Assert.NotSame(owner, window.ViewModel);
                owner.ErrorMessage = "The output folder is no longer available.";
                Assert.Equal("The output folder is no longer available.", window.ViewModel.ErrorMessage);
                Assert.Null(owner.ErrorMessage);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static Task CloseSelectedTabAsync(MainWindow window) => window.CloseSelectedDocumentTabAsync();

    private static string CreateDocument(string root, string name) {
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, name);
        PdfDocument.Create(document => document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(300)))).Save(path);
        return path;
    }
}

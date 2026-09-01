using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioDocumentTabHostTests {
    [Fact]
    public async Task TabsRetainIndependentLiveWorkspacesAndActivateExistingPaths() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-tabs-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string firstPath = Path.Combine(root, "first.pdf");
        string secondPath = Path.Combine(root, "second.pdf");
        CreateDocument(firstPath, 1);
        CreateDocument(secondPath, 2);
        MainWindowViewModel? activated = null;

        try {
            using var host = new StudioDocumentTabHost(
                openDocument => new MainWindowViewModel(
                    _ => Task.FromResult<string?>(null),
                    confirmUnsavedChanges: () => Task.FromResult(UnsavedChangesDecision.Discard),
                    openDocumentInTab: openDocument),
                document => activated = document);

            await host.OpenDocumentAsync(firstPath);
            StudioDocumentTabViewModel firstTab = Assert.Single(host.Tabs);
            Assert.Same(firstTab.Document, activated);
            firstTab.Document.SetOrganizerSelection([firstTab.Document.OrganizerPages[0]]);
            await firstTab.Document.DuplicateSelectedCommand.ExecuteAsync(null);
            Assert.True(firstTab.Document.IsDirty);
            Assert.Equal(2, firstTab.Document.Pages.Count);

            await host.OpenDocumentAsync(secondPath);
            Assert.Equal(2, host.Tabs.Count);
            StudioDocumentTabViewModel secondTab = host.SelectedTab!;
            Assert.Equal(2, secondTab.Document.Pages.Count);
            Assert.Same(secondTab.Document, activated);

            host.SelectedTab = firstTab;
            Assert.Same(firstTab.Document, activated);
            Assert.True(firstTab.Document.IsDirty);
            Assert.Equal(2, firstTab.Document.Pages.Count);

            await host.OpenDocumentAsync(secondPath);
            Assert.Equal(2, host.Tabs.Count);
            Assert.Same(secondTab, host.SelectedTab);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ClosingTabsUsesDocumentClosePolicyAndReturnsToEmptyWorkspace() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-tab-close-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string firstPath = Path.Combine(root, "first.pdf");
        string secondPath = Path.Combine(root, "second.pdf");
        CreateDocument(firstPath, 1);
        CreateDocument(secondPath, 1);
        MainWindowViewModel? activated = null;

        try {
            using var host = new StudioDocumentTabHost(
                openDocument => new MainWindowViewModel(
                    _ => Task.FromResult<string?>(null),
                    confirmUnsavedChanges: () => Task.FromResult(UnsavedChangesDecision.Discard),
                    openDocumentInTab: openDocument),
                document => activated = document);
            await host.OpenDocumentAsync(firstPath);
            await host.OpenDocumentAsync(secondPath);
            StudioDocumentTabViewModel firstTab = host.Tabs[0];
            StudioDocumentTabViewModel secondTab = host.Tabs[1];

            await host.CloseTabAsync(secondTab);
            Assert.Single(host.Tabs);
            Assert.Same(firstTab, host.SelectedTab);
            Assert.Same(firstTab.Document, activated);

            await host.CloseTabAsync(firstTab);
            Assert.Empty(host.Tabs);
            Assert.Null(host.SelectedTab);
            Assert.NotNull(activated);
            Assert.False(activated!.HasDocument);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task SaveAsRejectsAPathOwnedByAnotherLiveTab() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-tab-saveas-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string firstPath = Path.Combine(root, "first.pdf");
        string secondPath = Path.Combine(root, "second.pdf");
        CreateDocument(firstPath, 1);
        CreateDocument(secondPath, 2);
        StudioDocumentTabHost? host = null;

        try {
            host = new StudioDocumentTabHost(
                openDocument => new MainWindowViewModel(
                    _ => Task.FromResult<string?>(null),
                    pickSavePdf: _ => Task.FromResult<string?>(secondPath),
                    canSaveAsPath: path => host!.CanActiveDocumentOwnPath(path),
                    openDocumentInTab: openDocument),
                _ => { });
            await host.OpenDocumentAsync(firstPath);
            StudioDocumentTabViewModel firstTab = Assert.Single(host.Tabs);
            await host.OpenDocumentAsync(secondPath);
            host.SelectedTab = firstTab;

            await firstTab.Document.SaveAsCommand.ExecuteAsync(null);

            Assert.Equal(Path.GetFullPath(firstPath), firstTab.Document.DocumentPath);
            Assert.Contains("already open", firstTab.Document.OperationStatus, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(2, host.Tabs[1].Document.Pages.Count);
        } finally {
            host?.Dispose();
            Directory.Delete(root, recursive: true);
        }
    }

    private static void CreateDocument(string path, int pageCount) {
        PdfDocument.Create(compose => {
            for (int index = 0; index < pageCount; index++) {
                compose.Page(page => page.Size(420D, 620D));
            }
        }).Save(path);
    }
}

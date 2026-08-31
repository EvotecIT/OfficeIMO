using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Tests;

public sealed class MainWindowViewModelTests {
    [Fact]
    public async Task OpenCommandLoadsPathReturnedByPicker() {
        string fixture = GetFixturePath();
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(fixture));

        await viewModel.OpenCommand.ExecuteAsync(null);

        Assert.True(viewModel.HasDocument);
        Assert.False(viewModel.IsEmpty);
        Assert.Equal(System.IO.Path.GetFileName(fixture), viewModel.DocumentName);
        Assert.NotEmpty(viewModel.Pages);
        Assert.Same(viewModel.Pages[0], viewModel.SelectedPage);
    }

    [Fact]
    public async Task FailedOpenProducesDismissibleErrorState() {
        string missing = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            $"officeimo-studio-missing-{Guid.NewGuid():N}.pdf");
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(missing));

        await viewModel.OpenCommand.ExecuteAsync(null);

        Assert.False(viewModel.HasDocument);
        Assert.True(viewModel.IsEmpty);
        Assert.True(viewModel.HasError);
        Assert.Contains("no longer exists", viewModel.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        viewModel.DismissErrorCommand.Execute(null);
        Assert.False(viewModel.HasError);
    }

    [Fact]
    public void FitPageRecomputesWhenSelectedPageDimensionsChange() {
        using var coordinator = new PageRenderCoordinator((page, scale, _) =>
            Task.FromResult(new PdfRenderedPage(
                page,
                scale,
                [1],
                1,
                1,
                TimeSpan.Zero,
                Array.Empty<string>())));
        using var sceneCoordinator = new PageSceneCoordinator((page, _) =>
            Task.FromResult(TestPdfPageScenes.Create(page)));
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        var squarePage = new PdfPageViewModel(1, 400, 400, 0, 1D, sceneCoordinator, coordinator);
        var tallPage = new PdfPageViewModel(2, 400, 1000, 0, 1D, sceneCoordinator, coordinator);
        viewModel.Pages.Add(squarePage);
        viewModel.Pages.Add(tallPage);
        viewModel.SetViewportSize(1000, 700);
        viewModel.SelectedPage = squarePage;

        viewModel.FitPageCommand.Execute(null);
        double squareZoom = viewModel.Zoom;
        viewModel.SelectedPage = tallPage;

        Assert.True(viewModel.Zoom < squareZoom);
        Assert.Equal(0.63D, viewModel.Zoom);
    }

    [Fact]
    public async Task OrganizerMutationRefreshesReaderAndSupportsUndo() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-shell-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "editable.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(500, 700))).Save(path);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);
            viewModel.SetOrganizerSelection([viewModel.OrganizerPages[0]]);

            await viewModel.DuplicateSelectedCommand.ExecuteAsync(null);

            Assert.Equal(2, viewModel.Pages.Count);
            Assert.True(viewModel.IsDirty);
            Assert.True(viewModel.CanUndo);

            await viewModel.UndoCommand.ExecuteAsync(null);
            Assert.Single(viewModel.Pages);
            Assert.False(viewModel.IsDirty);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task BlockedOrganizerMutationDoesNotReportSuccessOrRefreshTheDocument() {
        string fixture = GetFixturePath();
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        await viewModel.OpenDocumentAsync(fixture);
        PdfPageViewModel originalPage = viewModel.Pages[0];
        viewModel.SetOrganizerSelection([viewModel.OrganizerPages[0]]);

        await viewModel.DuplicateSelectedCommand.ExecuteAsync(null);

        Assert.False(viewModel.CanMutatePages);
        Assert.Same(originalPage, viewModel.Pages[0]);
        Assert.Single(viewModel.Pages);
        Assert.Equal("Operation failed", viewModel.OperationStatus);
        Assert.True(viewModel.HasError);
    }

    [Fact]
    public async Task DiscardedDirtyDocumentDeletesRecoveryBeforeOpeningReplacement() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-transition-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string first = Path.Combine(root, "first.pdf");
        string second = Path.Combine(root, "second.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(500, 700))).Save(first);
        PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).Save(second);

        try {
            using (var viewModel = new MainWindowViewModel(
                _ => Task.FromResult<string?>(null),
                confirmUnsavedChanges: () => Task.FromResult(UnsavedChangesDecision.Discard))) {
                await viewModel.OpenDocumentAsync(first);
                viewModel.SetOrganizerSelection([viewModel.OrganizerPages[0]]);
                await viewModel.DuplicateSelectedCommand.ExecuteAsync(null);
                Assert.True(viewModel.IsDirty);

                await viewModel.OpenDocumentAsync(second);

                Assert.Equal("second.pdf", viewModel.DocumentName);
                Assert.Single(viewModel.Pages);
                Assert.False(viewModel.IsDirty);
            }

            using PdfWorkspace reopened = await PdfWorkspace.OpenAsync(first, CancellationToken.None);
            Assert.False(reopened.HasRecovery);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task DroppingOntoAnotherSelectedPageIsANoOp() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-drop-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "four-pages.pdf");
        PdfDocument.Create(compose => {
            for (int page = 0; page < 4; page++) compose.Page(item => item.Size(500 + page, 700));
        }).Save(path);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);
            PdfPageViewModel[] original = viewModel.Pages.ToArray();
            viewModel.SetOrganizerSelection([viewModel.OrganizerPages[1], viewModel.OrganizerPages[2]]);

            await viewModel.ReorderByDropAsync(2, 3);

            Assert.False(viewModel.IsDirty);
            Assert.Equal(original, viewModel.Pages);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task StandardNamedPageActionsNavigateInsideTheDocument() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-links-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "links.pdf");
        PdfDocument.Create(compose => {
            compose.Page(page => page.Size(500, 700));
            compose.Page(page => page.Size(500, 700));
        }).Save(path);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);

            await viewModel.ActivatePageLinkAsync("LastPage");
            Assert.Equal(2, viewModel.SelectedPage?.PageNumber);
            await viewModel.ActivatePageLinkAsync("PrevPage");
            Assert.Equal(1, viewModel.SelectedPage?.PageNumber);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task CancellingDirtyDocumentTransitionKeepsTheCurrentWorkspace() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-transition-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string first = Path.Combine(root, "first.pdf");
        string second = Path.Combine(root, "second.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(500, 700))).Save(first);
        PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).Save(second);
        UnsavedChangesDecision decision = UnsavedChangesDecision.Cancel;

        try {
            using var viewModel = new MainWindowViewModel(
                _ => Task.FromResult<string?>(null),
                confirmUnsavedChanges: () => Task.FromResult(decision));
            await viewModel.OpenDocumentAsync(first);
            viewModel.SetOrganizerSelection([viewModel.OrganizerPages[0]]);
            await viewModel.DuplicateSelectedCommand.ExecuteAsync(null);

            await viewModel.OpenDocumentAsync(second);

            Assert.Equal("first.pdf *", viewModel.DocumentName);
            Assert.Equal(2, viewModel.Pages.Count);
            Assert.True(viewModel.IsDirty);

            decision = UnsavedChangesDecision.Discard;
            Assert.True(await viewModel.RequestCloseDocumentAsync());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ClosingDocumentClearsDocumentScopedSearchAndStatus() {
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        await viewModel.OpenDocumentAsync(GetFixturePath());
        viewModel.SearchQuery = "test";
        await viewModel.SearchCommand.ExecuteAsync(null);
        Assert.NotEmpty(viewModel.SearchResults);
        Assert.True(viewModel.HasOperationStatus);

        Assert.True(await viewModel.RequestCloseDocumentAsync());

        Assert.False(viewModel.HasDocument);
        Assert.Empty(viewModel.SearchQuery);
        Assert.Empty(viewModel.SearchResults);
        Assert.Null(viewModel.SelectedSearchResult);
        Assert.False(viewModel.HasOperationStatus);
        Assert.Equal(0D, viewModel.OperationProgressFraction);
    }

    private static string GetFixturePath() =>
        System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf");
}

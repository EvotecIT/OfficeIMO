using OfficeIMO.Pdf;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class ReaderLayoutViewModelTests {
    [Fact]
    public async Task LiveReaderBindingsPreserveSelectionAndVirtualizeGridRows() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-reader-live-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "many-pages.pdf");
        CreateDocument(source, pageCount: 40);

        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var window = new MainWindow();
                try {
                    window.Show();
                    window.Width = 1280;
                    window.Height = 820;
                    await window.TabHost.OpenDocumentAsync(source);
                    window.Measure(new Size(1280, 820));
                    window.Arrange(new Rect(0, 0, 1280, 820));
                    MainWindowViewModel viewModel = window.ViewModel;

                    viewModel.SelectedReaderLayoutChoice = GetLayout(viewModel, ReaderLayoutMode.SinglePage);
                    viewModel.NextPageCommand.Execute(null);
                    Assert.Equal(2, viewModel.SelectedPage?.PageNumber);
                    Assert.Same(viewModel.SelectedPage, window.ReaderPagesListControl.SelectedItem);

                    viewModel.SelectedReaderLayoutChoice = GetLayout(viewModel, ReaderLayoutMode.TwoPage);
                    viewModel.NextPageCommand.Execute(null);
                    Assert.Equal(3, viewModel.SelectedPage?.PageNumber);
                    Assert.Equal([2, 3], viewModel.ReaderPages.Select(page => page.PageNumber));
                    Assert.Same(viewModel.SelectedPage, window.ReaderPagesListControl.SelectedItem);

                    viewModel.SelectedReaderLayoutChoice = GetLayout(viewModel, ReaderLayoutMode.Grid);
                    window.Measure(new Size(1280, 820));
                    window.Arrange(new Rect(0, 0, 1280, 820));
                    using var renderedFrame = window.CaptureRenderedFrame();
                    Assert.Empty(viewModel.ReaderPages);
                    Assert.True(viewModel.ReaderGridRows.Count > 1);
                    int realizedPages = window.ReaderGridPagesListControl
                        .GetVisualDescendants()
                        .OfType<PdfPageView>()
                        .Count();
                    Assert.InRange(realizedPages, 1, 16);
                } finally {
                    window.Close();
                }

                return true;
            }, CancellationToken.None);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task LayoutsReuseTheActiveDocumentPagesAndBuildCoverAwareSpreads() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-reader-layout-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "five-pages.pdf");
        CreateDocument(source, pageCount: 5);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(source);

            Assert.Equal(ReaderLayoutMode.Continuous, viewModel.ReaderLayout);
            Assert.True(viewModel.ReaderPages.SequenceEqual(viewModel.Pages));

            viewModel.SelectedReaderLayoutChoice = GetLayout(viewModel, ReaderLayoutMode.SinglePage);
            viewModel.SelectedPage = viewModel.Pages[3];
            Assert.Equal([4], viewModel.ReaderPages.Select(page => page.PageNumber));

            viewModel.SelectedReaderLayoutChoice = GetLayout(viewModel, ReaderLayoutMode.TwoPage);
            Assert.Equal([4, 5], viewModel.ReaderPages.Select(page => page.PageNumber));

            viewModel.SelectedPage = viewModel.Pages[2];
            Assert.Equal([2, 3], viewModel.ReaderPages.Select(page => page.PageNumber));

            viewModel.SelectedPage = viewModel.Pages[0];
            Assert.Equal([1], viewModel.ReaderPages.Select(page => page.PageNumber));

            viewModel.SelectedReaderLayoutChoice = GetLayout(viewModel, ReaderLayoutMode.Grid);
            Assert.Empty(viewModel.ReaderPages);
            Assert.True(viewModel.ReaderGridRows.SelectMany(row => row.Pages).SequenceEqual(viewModel.Pages));
            Assert.All(viewModel.ReaderGridRows, row => Assert.InRange(row.Pages.Count, 1, 4));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ReaderCommandsApplyNightModeAndNavigateDocumentBounds() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-reader-commands-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "three-pages.pdf");
        CreateDocument(source, pageCount: 3);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(source);

            viewModel.LastPageCommand.Execute(null);
            Assert.Equal(3, viewModel.SelectedPage?.PageNumber);
            viewModel.FirstPageCommand.Execute(null);
            Assert.Equal(1, viewModel.SelectedPage?.PageNumber);

            viewModel.TogglePageNightModeCommand.Execute(null);
            Assert.True(viewModel.IsPageNightMode);
            Assert.All(viewModel.Pages, page => Assert.True(page.IsNightMode));

            viewModel.TogglePageNightModeCommand.Execute(null);
            Assert.False(viewModel.IsPageNightMode);
            Assert.All(viewModel.Pages, page => Assert.False(page.IsNightMode));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task TwoPageFitAccountsForTheWholeSpread() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-reader-fit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "spread.pdf");
        CreateDocument(source, pageCount: 3);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(source);
            viewModel.SetViewportSize(1000D, 1000D);
            viewModel.SelectedPage = viewModel.Pages[1];

            viewModel.SelectedReaderLayoutChoice = GetLayout(viewModel, ReaderLayoutMode.SinglePage);
            double singlePageZoom = viewModel.Zoom;
            viewModel.SelectedReaderLayoutChoice = GetLayout(viewModel, ReaderLayoutMode.TwoPage);

            Assert.Equal([2, 3], viewModel.ReaderPages.Select(page => page.PageNumber));
            Assert.True(viewModel.Zoom < singlePageZoom);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static ReaderLayoutChoice GetLayout(MainWindowViewModel viewModel, ReaderLayoutMode mode) =>
        viewModel.ReaderLayoutChoices.Single(choice => choice.Mode == mode);

    private static void CreateDocument(string path, int pageCount) {
        PdfDocument.Create(compose => {
            for (int index = 0; index < pageCount; index++) {
                compose.Page(page => page.Size(420D, 620D));
            }
        }).Save(path);
    }
}

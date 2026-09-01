using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class ReaderComparisonViewModelTests {
    [Fact]
    public async Task ComparisonUsesIndependentReadOnlyPagesAndSynchronizesNavigationZoomAndNightMode() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-comparison-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string primaryPath = Path.Combine(root, "primary.pdf");
        string comparisonPath = Path.Combine(root, "comparison.pdf");
        CreateDocument(primaryPath, 3, 420D);
        CreateDocument(comparisonPath, 2, 500D);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(comparisonPath));
            await viewModel.OpenDocumentAsync(primaryPath);
            ReaderLayoutChoice originalLayout = viewModel.SelectedReaderLayoutChoice;

            await viewModel.OpenComparisonCommand.ExecuteAsync(null);

            Assert.True(viewModel.IsComparisonOpen);
            Assert.Equal("comparison.pdf", viewModel.ComparisonDocumentName);
            Assert.Equal(2, viewModel.ComparisonPages.Count);
            Assert.Equal(ReaderLayoutMode.SinglePage, viewModel.ReaderLayout);
            Assert.Equal(1, viewModel.PrimaryReaderColumnSpan);
            Assert.Single(viewModel.ComparisonReaderPages);

            viewModel.SelectedPage = viewModel.Pages[2];
            Assert.Equal(2, viewModel.ComparisonSelectedPage?.PageNumber);
            Assert.Equal([2], viewModel.ComparisonReaderPages.Select(page => page.PageNumber));

            viewModel.ComparisonSelectedPage = viewModel.ComparisonPages[0];
            Assert.Equal(1, viewModel.SelectedPage?.PageNumber);

            await viewModel.ActivateComparisonPageLinkAsync("LastPage");
            Assert.Equal(2, viewModel.ComparisonSelectedPage?.PageNumber);
            Assert.Equal(2, viewModel.SelectedPage?.PageNumber);

            viewModel.ActualSizeCommand.Execute(null);
            Assert.Equal(420D, viewModel.Pages[0].DisplayWidth);
            Assert.Equal(500D, viewModel.ComparisonPages[0].DisplayWidth);
            viewModel.ZoomInCommand.Execute(null);
            Assert.Equal(525D, viewModel.Pages[0].DisplayWidth);
            Assert.Equal(625D, viewModel.ComparisonPages[0].DisplayWidth);

            viewModel.TogglePageNightModeCommand.Execute(null);
            Assert.All(viewModel.Pages, page => Assert.True(page.IsNightMode));
            Assert.All(viewModel.ComparisonPages, page => Assert.True(page.IsNightMode));

            viewModel.CloseComparisonCommand.Execute(null);
            Assert.False(viewModel.IsComparisonOpen);
            Assert.Empty(viewModel.ComparisonPages);
            Assert.Empty(viewModel.ComparisonReaderPages);
            Assert.Equal(3, viewModel.PrimaryReaderColumnSpan);
            Assert.Equal(originalLayout, viewModel.SelectedReaderLayoutChoice);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ComparisonRejectsTheActiveDocumentPath() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-comparison-same-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "document.pdf");
        CreateDocument(path, 1, 420D);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(path));
            await viewModel.OpenDocumentAsync(path);

            await viewModel.OpenComparisonCommand.ExecuteAsync(null);

            Assert.False(viewModel.IsComparisonOpen);
            Assert.Empty(viewModel.ComparisonPages);
            Assert.Contains("different PDF", viewModel.OperationStatus, StringComparison.Ordinal);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static void CreateDocument(string path, int pageCount, double width) {
        PdfDocument.Create(compose => {
            for (int index = 0; index < pageCount; index++) {
                compose.Page(page => page.Size(width, 620D));
            }
        }).Save(path);
    }
}

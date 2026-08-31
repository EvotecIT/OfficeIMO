using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

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
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        var squarePage = new PdfPageViewModel(1, 400, 400, 0, 1D, coordinator);
        var tallPage = new PdfPageViewModel(2, 400, 1000, 0, 1D, coordinator);
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

    private static string GetFixturePath() =>
        System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf");
}

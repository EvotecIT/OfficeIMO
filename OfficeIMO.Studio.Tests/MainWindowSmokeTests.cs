using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class MainWindowSmokeTests {
    [Fact]
    public async Task CreatesAndLaysOutResponsiveStudioSurfaces() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var window = new MainWindow();
            try {
                window.Show();
                window.Measure(new Avalonia.Size(1280, 820));
                window.Arrange(new Avalonia.Rect(0, 0, 1280, 820));

                Assert.NotNull(window.DataContext);
                Assert.IsType<MainWindowViewModel>(window.DataContext);
                Assert.False(window.ViewModel.HasDocument);
                Assert.True(window.ViewModel.IsEmpty);
                Assert.True(window.ViewModel.IsHomeMode);

                window.ViewModel.ShowPdfWorkspaceCommand.Execute(null);
                window.Width = 1050;
                window.Height = 560;
                window.Measure(new Avalonia.Size(1050, 560));
                window.Arrange(new Avalonia.Rect(0, 0, 1050, 560));
                window.ApplyResponsiveLayout(1050);
                Assert.True(window.ViewModel.IsPdfWorkspaceMode);
                Assert.True(window.IsCompactLayout);
                Assert.False(window.AreFitShortcutsVisible);

                window.ViewModel.ShowConversionWorkbenchCommand.Execute(null);
                window.Measure(new Avalonia.Size(1050, 560));
                window.Arrange(new Avalonia.Rect(0, 0, 1050, 560));
                Assert.True(window.ViewModel.IsConversionMode);
                Assert.True(window.IsConversionCompact);

                window.ViewModel.ShowDocumentHealthCommand.Execute(null);
                window.Measure(new Avalonia.Size(1050, 560));
                window.Arrange(new Avalonia.Rect(0, 0, 1050, 560));
                Assert.True(window.IsDocumentHealthCompact);
                window.Width = 1600;
                window.Height = 1000;
                window.Measure(new Avalonia.Size(1600, 1000));
                window.Arrange(new Avalonia.Rect(0, 0, 1600, 1000));
                window.ApplyResponsiveLayout(1600);
                Assert.True(window.ViewModel.IsDocumentHealthMode);
                Assert.False(window.IsCompactLayout);
                Assert.False(window.IsDocumentHealthCompact);
                Assert.True(window.AreFitShortcutsVisible);
            } finally {
                window.Close();
            }

            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData("notes.txt", "report.PDF", "report.PDF")]
    [InlineData("first.pdf", "second.pdf", "first.pdf")]
    public void DropSelectionUsesFirstPdfCaseInsensitively(
        string first,
        string second,
        string expected) {
        bool found = MainWindow.TryGetPdfPath([first, second], out string? path);

        Assert.True(found);
        Assert.Equal(expected, path);
    }

    [Fact]
    public void DropSelectionRejectsMissingOrUnsupportedFiles() {
        Assert.False(MainWindow.TryGetPdfPath([null, "notes.txt", "image.png"], out string? path));
        Assert.Null(path);
    }
}

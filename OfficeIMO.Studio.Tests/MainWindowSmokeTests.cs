using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class MainWindowSmokeTests {
    [Fact]
    public async Task CreatesAndLaysOutEmptyReaderShell() {
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

                window.ViewModel.ShowConversionWorkbenchCommand.Execute(null);
                window.Measure(new Avalonia.Size(1280, 820));
                window.Arrange(new Avalonia.Rect(0, 0, 1280, 820));
                Assert.True(window.ViewModel.IsConversionMode);

                window.ViewModel.ShowDocumentHealthCommand.Execute(null);
                window.Measure(new Avalonia.Size(1280, 820));
                window.Arrange(new Avalonia.Rect(0, 0, 1280, 820));
                Assert.True(window.ViewModel.IsDocumentHealthMode);
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

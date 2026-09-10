using System.Reflection;
using Avalonia.Platform.Storage;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class PrinterOutputPickerTests {
    [Fact]
    public async Task PrintWorkbenchUsesThePrinterPickerInsteadOfThePdfSavePicker() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSavePdf: _ => throw new InvalidOperationException("The PDF picker must not handle printer output."),
                pickPrintOutput: _ => Task.FromResult<string?>("selected.xps"));
            await model.OutputWorkbench.PrintPreview.ChoosePrintOutputCommand.ExecuteAsync(null);
            Assert.Equal("selected.xps", model.OutputWorkbench.PrintPreview.PrintOutputPath);
            return true;
        }, default);
    }

    [Theory]
    [InlineData("output.xps")]
    [InlineData("output.prn")]
    [InlineData("output.pdf")]
    public async Task PrinterPickerPreservesDriverOutputExtension(string outputName) {
        string path = Path.Combine(Path.GetTempPath(), outputName);
        var file = new TestStorageFile(new Uri(path).AbsoluteUri, [], outputName);
        var provider = DispatchProxy.Create<IStorageProvider, TestStorageFile.StorageProxy>();
        FilePickerSaveOptions? options = null;
        ((TestStorageFile.StorageProxy)(object)provider).Call = (method, arguments) => method switch {
            "get_CanSave" => true,
            "SaveFilePickerAsync" => Capture((FilePickerSaveOptions)arguments![0]!),
            _ => throw new NotSupportedException(method)
        };
        Task<IStorageFile?> Capture(FilePickerSaveOptions value) {
            options = value;
            return Task.FromResult<IStorageFile?>(file.Item);
        }
        Assert.Equal(path, await MainWindow.PickPrintOutputAsync(provider, "Printer output", "source.pdf *", default));
        Assert.NotNull(options);
        Assert.Null(options.DefaultExtension);
        Assert.Equal("source", options.SuggestedFileName);
        Assert.Contains(Assert.Single(options.FileTypeChoices!).Patterns!, pattern =>
            System.IO.Enumeration.FileSystemName.MatchesSimpleExpression(pattern, outputName, ignoreCase: true));
        Assert.Equal(1, file.Disposals);
        Assert.Equal(0, file.Writes);
    }
}

using Avalonia.Platform.Storage;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow {
    private Task<string?> PickPrintOutputAsync(CancellationToken token) => PickPrintOutputAsync(StorageProvider,
        _services.Localizer.GetOrDefault("PrintPreview.ChooseOutput", "Choose printer output"), ViewModel.DocumentName, token);

    internal static async Task<string?> PickPrintOutputAsync(IStorageProvider provider, string title, string documentName, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (!provider.CanSave) return null;
        using IStorageFile? file = await provider.SaveFilePickerAsync(new FilePickerSaveOptions {
            Title = title,
            SuggestedFileName = Path.GetFileNameWithoutExtension(documentName.TrimEnd(' ', '*')),
            FileTypeChoices = [FilePickerFileTypes.All]
        });
        token.ThrowIfCancellationRequested();
        if (file is null) return null;
        return file.TryGetLocalPath() ?? throw new IOException("File printers require a local output path.");
    }
}

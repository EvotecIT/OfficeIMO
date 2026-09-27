using Avalonia.Platform.Storage;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Shell;

public partial class MainWindow {
    private IStudioFileDialogs? _fileDialogs;

    private IStudioFileDialogs FileDialogs => _fileDialogs ??= new WindowFileDialogs(this);

    // Pickers register every chosen location with Studio storage, so provider-backed files work like local ones.
    private sealed class WindowFileDialogs(MainWindow owner) : IStudioFileDialogs {
        public Task<string?> PickOpenFileAsync(string title, StudioFileType type, CancellationToken cancellationToken) =>
            owner.PickFileSafelyAsync(async token => {
                if (!owner.StorageProvider.CanOpen) return null;
                var files = await owner.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
                    Title = title, AllowMultiple = false, FileTypeFilter = [ToFilter(type)]
                });
                return await owner._services.Storage.RegisterSingleAsync(files, token).ConfigureAwait(true);
            }, cancellationToken);

        public Task<string?> PickSaveFileAsync(string title, string suggestedName, StudioFileType type, CancellationToken cancellationToken) =>
            owner.PickFileSafelyAsync(async token => {
                if (!owner.StorageProvider.CanSave) return null;
                string? extension = type.Extensions.FirstOrDefault(value => value != "*");
                IStorageFile? file = await owner.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions {
                    Title = title, SuggestedFileName = suggestedName, DefaultExtension = extension,
                    FileTypeChoices = [ToFilter(type)]
                });
                string? location = await owner._services.Storage.RegisterSingleAsync(file is null ? [] : [file], token).ConfigureAwait(true);
                return location is not null && await owner.ConfirmProviderWriteAsync(location) ? location : null;
            }, cancellationToken);

        public Task<string?> PickFolderAsync(string title, CancellationToken cancellationToken) =>
            owner.PickFileSafelyAsync(async token => {
                if (!owner.StorageProvider.CanPickFolder) return null;
                var folders = await owner.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions { Title = title, AllowMultiple = false });
                return await owner._services.Storage.RegisterFolderAsync(folders, token).ConfigureAwait(true);
            }, cancellationToken);

        private static FilePickerFileType ToFilter(StudioFileType type) => new(type.Name) {
            Patterns = type.Extensions.Select(extension => extension == "*" ? "*" : "*." + extension).ToArray(),
            MimeTypes = type.MimeType is null ? null : [type.MimeType]
        };
    }
}

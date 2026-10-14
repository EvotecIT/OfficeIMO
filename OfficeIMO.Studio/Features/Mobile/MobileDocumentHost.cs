using Avalonia.Controls;
using Avalonia.Platform.Storage;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Mobile;

/// <summary>Connects shared document commands to the mobile presentation and permission-scoped storage.</summary>
internal sealed partial class MobileDocumentHost(
    StudioApplicationServices services, MobileWorkspaceView workspace, Func<string, Task> share, Func<IStorageProvider> provider) : IStudioFileDialogs {
    private IStorageProvider Provider => provider();

    internal Task<T?> ShowAsync<T>(StudioDialogContent content) => workspace.ShowDialogAsync<T>(content);

    public async Task<string?> PickOpenFileAsync(string title, StudioFileType type, CancellationToken cancellationToken) =>
        (await PickFilesAsync(title, type, false, cancellationToken)).FirstOrDefault();

    internal async Task<IReadOnlyList<string>> PickFilesAsync(string title, StudioFileType type, bool multiple, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var files = await Provider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = title, AllowMultiple = multiple, FileTypeFilter = [ToFilter(type)]
        });
        return await services.Storage.RegisterManyAsync(files, token);
    }

    public async Task<string?> PickSaveFileAsync(string title, string suggestedName, StudioFileType type, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        // UIKit's export picker may publish a placeholder while selecting the destination. Obtain consent first.
        if (!await ShowAsync<bool>(new ProviderSaveDialogContent(suggestedName, services.Localizer))) return null;
        cancellationToken.ThrowIfCancellationRequested();
        IStorageFile? file = await Provider.SaveFilePickerAsync(new FilePickerSaveOptions {
            Title = title, SuggestedFileName = suggestedName,
            DefaultExtension = type.Extensions.FirstOrDefault(value => value != "*"),
            FileTypeChoices = [ToFilter(type)]
        });
        string? location = await services.Storage.RegisterSingleAsync(file is null ? [] : [file], cancellationToken);
        return location;
    }

    public async Task<string?> PickFolderAsync(string title, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var folders = await Provider.OpenFolderPickerAsync(new FolderPickerOpenOptions { Title = title, AllowMultiple = false });
        return await services.Storage.RegisterFolderAsync(folders, cancellationToken);
    }

    internal async Task<string?> PickWorkingCopyAsync(string title, StudioFileType type, CancellationToken token) {
        string? location = await PickOpenFileAsync(title, type, token);
        if (location is null) return null;
        var root = services.LocalDocuments ?? throw new IOException(Infrastructure.Localization.StudioLocalization.Current.Get("Mobile.DocumentFolderUnavailable"));
        StudioStorageSnapshot source = await services.Storage.ReadSnapshotAsync(location, token);
        return await root.WriteWorkingCopyAsync(services.Storage.Describe(location).Name, source.Bytes, token);
    }

    internal async Task<byte[]?> PickImageAsync(CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var files = await Provider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = services.Localizer.Get("Picker.ChooseImage"), AllowMultiple = false,
            FileTypeFilter = [ToFilter(new("Images", ["png", "jpg", "jpeg"]))]
        });
        return await StudioStorageInput.ReadImageAsync(files, token);
    }

    internal Task<bool> ConfirmWriteAsync(string location, bool workflow = false) =>
        !services.Storage.UsesProviderPublication(location) ? Task.FromResult(true) :
            ShowAsync<bool>(new ProviderSaveDialogContent(services.Storage.Describe(location).Name,
                services.Localizer, workflow, services.Storage.IsFolder(location)));

    internal async Task<string?> PromptPasswordAsync(string name, bool invalid, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        string? password = await ShowAsync<string>(new PdfPasswordDialogContent(name, invalid, services.Localizer));
        token.ThrowIfCancellationRequested();
        return password;
    }

    internal async Task OpenUriAsync(Uri uri) {
        string location = uri.IsFile ? uri.LocalPath : uri.AbsoluteUri;
        if (services.Storage.IsFolder(location) || uri.IsFile && Directory.Exists(location))
            throw new IOException(Infrastructure.Localization.StudioLocalization.Current.Get("Mobile.BrowseOutputFolder"));
        if (uri.IsFile || services.Storage.IsKnownLocation(location)) {
            await MobileFileSharing.ShareAsync(services, location, share);
            return;
        }
        var topLevel = TopLevel.GetTopLevel(workspace) ?? throw new IOException(Infrastructure.Localization.StudioLocalization.Current.Get("Mobile.NotReadyToOpenLink"));
        if (!await topLevel.Launcher.LaunchUriAsync(uri)) throw new IOException(services.Localizer.Get("Error.CouldNotOpenLink"));
    }

    private static FilePickerFileType ToFilter(StudioFileType type) => new(type.Name) {
        Patterns = type.Extensions.Select(extension => extension == "*" ? "*" : "*." + extension.TrimStart('.')).ToArray(),
        MimeTypes = type.MimeType is null ? null : [type.MimeType],
        AppleUniformTypeIdentifiers = type.Extensions.Contains("*") ? ["public.item"] : null
    };
}

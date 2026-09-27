using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Open and Show in folder for any finished output location, shared by every result screen.</summary>
public sealed partial class StudioOutputActions {
    private readonly Func<string, CancellationToken, Task> _open;
    private readonly Func<string, Task> _revealFolder;
    private readonly Action<string> _reportError;

    internal StudioOutputActions(Func<string, CancellationToken, Task> open, Func<string, Task> revealFolder, Action<string> reportError) {
        _open = open;
        _revealFolder = revealFolder;
        _reportError = reportError;
    }

    private static bool HasLocation(string? location) => !string.IsNullOrWhiteSpace(location);

    private static bool IsLocal(string? location) => HasLocation(location) && OfficeStorageIdentity.GetLocalPath(location!) is not null;

    [RelayCommand(CanExecute = nameof(HasLocation))]
    private async Task OpenAsync(string? location, CancellationToken cancellationToken) {
        if (!HasLocation(location)) return;
        try {
            if (OfficeStorageIdentity.GetLocalPath(location!) is { } local && Directory.Exists(local)) await _revealFolder(local).ConfigureAwait(true);
            else await _open(location!, cancellationToken).ConfigureAwait(true);
        } catch (Exception error) when (error is not OperationCanceledException) { _reportError(error.Message); }
    }

    [RelayCommand(CanExecute = nameof(IsLocal))]
    private async Task RevealAsync(string? location) {
        if (OfficeStorageIdentity.GetLocalPath(location ?? string.Empty) is not { } local) return;
        string folder = Directory.Exists(local) ? local : Path.GetDirectoryName(local) ?? local;
        try { await _revealFolder(folder).ConfigureAwait(true); }
        catch (Exception error) { _reportError(error.Message); }
    }
}

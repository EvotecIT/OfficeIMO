using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    /// <summary>Whether the selected pages can have their saved rotation baked into content and geometry.</summary>
    public bool CanNormalizePageRotation => HasOrganizerSelection && _workspace?.CanNormalizePageRotation == true;

    [RelayCommand]
    private async Task NormalizePageRotationAsync(CancellationToken cancellationToken) {
        int[] pages = GetSelectedPages();
        if (_workspace is null || pages.Length == 0 || !CanNormalizePageRotation) return;
        await RunMutationAsync(token => _workspace.NormalizePageRotationAsync(pages, token, CreateProgress()),
            cancellationToken, pages).ConfigureAwait(true);
    }
}

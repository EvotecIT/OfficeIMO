using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private int _viewRotationDegrees;
    /// <summary>Reader presentation angle; changing it does not mutate the document or its edit history.</summary>
    public int ViewRotationDegrees => _viewRotationDegrees;
    public Avalonia.Media.ITransform ViewRotationTransform => new Avalonia.Media.RotateTransform(_viewRotationDegrees);

    [RelayCommand] private void RotateViewLeft() => SetViewRotation(_viewRotationDegrees - 90);
    [RelayCommand] private void RotateViewRight() => SetViewRotation(_viewRotationDegrees + 90);
    [RelayCommand] private void ResetViewRotation() => SetViewRotation(0);

    private void SetViewRotation(int degrees, bool saveState = true) {
        _viewRotationDegrees = ((degrees % 360) + 360) % 360;
        foreach (var page in Pages) page.SetViewRotation(_viewRotationDegrees);
        foreach (var page in ComparisonPages) page.SetViewRotation(_viewRotationDegrees);
        ApplyFitZoom();
        OnPropertyChanged(nameof(ViewRotationDegrees));
        OnPropertyChanged(nameof(ViewRotationTransform));
        if (saveState) SaveDocumentViewState();
    }
}

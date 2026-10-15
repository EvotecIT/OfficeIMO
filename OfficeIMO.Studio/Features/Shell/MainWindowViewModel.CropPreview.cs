using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Organizer;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>
/// Cropping follows the inspector's pending-change model: changing the margin outlines the result on the selected pages,
/// Crop applies it as one undoable step, and Cancel returns the margin to its default and removes the outline.
/// </summary>
public sealed partial class MainWindowViewModel {
    internal const double DefaultCropMargin = 12D;
    private bool _cropPending;

    /// <summary>True while a changed crop margin is outlined but not yet applied.</summary>
    public bool HasPendingCrop => _cropPending;

    partial void OnCropMarginChanged(double value) {
        _cropPending = true;
        UpdateCropPreview();
    }

    [RelayCommand]
    private void CancelCrop() {
        CropMargin = DefaultCropMargin;
        EndCropPreview();
    }

    private void EndCropPreview() {
        _cropPending = false;
        UpdateCropPreview();
    }

    private void UpdateCropPreview() {
        bool show = _cropPending && IsPagesDocumentMode && CropMargin > 0D;
        foreach (PdfOrganizerPageViewModel page in OrganizerPages) page.CropPreviewMargin = show && page.IsSelected ? CropMargin : null;
        OnPropertyChanged(nameof(HasPendingCrop));
    }
}

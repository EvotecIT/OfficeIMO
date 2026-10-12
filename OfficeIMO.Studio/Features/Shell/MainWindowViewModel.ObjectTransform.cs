using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    [ObservableProperty] private bool _placeEditedImageBehindContent;
    [ObservableProperty] private double _objectScalePercent = 100;

    /// <summary>True while a changed size for the selected image is outlined on the page but not yet applied.</summary>
    public bool HasPendingImageScale => SelectedObject is { Kind: PdfEditorSelectionKind.Image } && ObjectScalePercent > 0 && Math.Abs(ObjectScalePercent - 100D) > 0.01D;

    partial void OnObjectScalePercentChanged(double value) => UpdateObjectPreview();

    // The image keeps its centre when resized, so the outline shows exactly where Apply would put it.
    private void UpdateObjectPreview() {
        Avalonia.Rect? preview = null;
        int? pageNumber = null;
        if (HasPendingImageScale && SelectedObject is { } selection) {
            double scale = ObjectScalePercent / 100D;
            PdfEditorVisualBounds bounds = selection.Bounds;
            double width = bounds.Width * scale, height = bounds.Height * scale;
            preview = new Avalonia.Rect((bounds.Left + bounds.Right - width) / 2D, (bounds.Top + bounds.Bottom - height) / 2D, width, height);
            pageNumber = selection.PageNumber;
        }
        foreach (PdfPageViewModel page in Pages) page.PendingObjectPreview = page.PageNumber == pageNumber ? preview : null;
        OnPropertyChanged(nameof(HasPendingImageScale));
    }

    [RelayCommand]
    private void CancelImageScale() => ObjectScalePercent = 100;

    private async void OnPageObjectTransform(PdfObjectTransformGesture gesture) {
        bool currentSelection = gesture.Selection.Kind == PdfEditorSelectionKind.Annotation
            ? Pages.Any(page => ReferenceEquals(page.SelectedObject, gesture.Selection))
            : ReferenceEquals(SelectedObject, gesture.Selection);
        if (_workspace is not { } workspace || IsWorkspaceBusy || !currentSelection ||
            (gesture.Selection.Kind == PdfEditorSelectionKind.Image ? !CanEditPageContent : !CanEditAnnotations)) return;
        long revision = workspace.Revision;
        PdfImageEditLayer layer = PlaceEditedImageBehindContent ? PdfImageEditLayer.BehindExistingContent : PdfImageEditLayer.AboveExistingContent;
        ClearObjectSelection();
        await RunMutationAsync(token => workspace.TransformSelectedObjectAsync(gesture, revision, layer, token, CreateProgress()), CancellationToken.None).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task ScaleSelectedImageAsync(CancellationToken token) {
        if (_workspace is not { } workspace || SelectedObject is not { Kind: PdfEditorSelectionKind.Image } selection || !CanEditPageContent) return;
        double scale = ObjectScalePercent / 100;
        if (!double.IsFinite(scale) || scale <= 0) return;
        PdfEditorVisualBounds bounds = selection.Bounds;
        double centerX = (bounds.Left + bounds.Right) / 2, centerY = (bounds.Top + bounds.Bottom) / 2;
        var target = new PdfEditorVisualBounds(centerX - bounds.Width * scale / 2, centerY - bounds.Height * scale / 2,
            centerX + bounds.Width * scale / 2, centerY + bounds.Height * scale / 2);
        long revision = workspace.Revision;
        PdfImageEditLayer layer = PlaceEditedImageBehindContent ? PdfImageEditLayer.BehindExistingContent : PdfImageEditLayer.AboveExistingContent;
        ClearObjectSelection();
        await RunMutationAsync(cancellation => workspace.TransformSelectedObjectAsync(new(selection, target), revision, layer,
            cancellation, CreateProgress()), token).ConfigureAwait(true);
    }
}

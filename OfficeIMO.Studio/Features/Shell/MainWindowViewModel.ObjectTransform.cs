using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    [ObservableProperty] private bool _placeEditedImageBehindContent;
    [ObservableProperty] private double _objectScalePercent = 100;

    private async void OnPageObjectTransform(PdfObjectTransformGesture gesture) {
        if (_workspace is not { } workspace || IsWorkspaceBusy || !ReferenceEquals(SelectedObject, gesture.Selection) ||
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

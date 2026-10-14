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
        try {
            Guid[] annotationIdentities = gesture.KeepSelection && gesture.Selection.Kind == PdfEditorSelectionKind.Annotation
                ? (gesture.Annotations is { Count: > 0 } annotations ? annotations : [gesture.Selection])
                    .Where(item => item.ObjectNumber.HasValue).Select(item => workspace.GetAnnotationIdentity(item.ObjectNumber!.Value)).ToArray()
                : [];
            ClearObjectSelection();
            bool applied = await RunMutationAsync(token => workspace.TransformSelectedObjectAsync(gesture, revision, layer, token, CreateProgress()), CancellationToken.None).ConfigureAwait(true);
            if (!applied || !gesture.KeepSelection || !ReferenceEquals(_workspace, workspace)) return;
            if (annotationIdentities.Length > 0) {
                var selected = workspace.AnnotationMetadata.Where(annotation => annotation.ObjectNumber is int number &&
                    annotationIdentities.Contains(workspace.GetAnnotationIdentity(number))).Select(annotation =>
                        new PdfEditorSelection(PdfEditorSelectionKind.Annotation, annotation.PageNumber!.Value, default,
                            ObjectNumber: annotation.ObjectNumber, Subtype: annotation.Subtype)).ToArray();
                OnPageAnnotationsSelected(new(selected, false));
            } else {
                // Image placements belong to a revision. Reacquire the transformed placement before the next key press.
                var matches = workspace.CreateDocumentSnapshot().Render.Interactions(gesture.Selection.PageNumber).Regions
                    .Where(region => region.Kind == PdfInteractionKind.Image &&
                        Math.Abs(region.Quad.Left - gesture.Target.Left) < 0.01 && Math.Abs(region.Quad.Top - gesture.Target.Top) < 0.01 &&
                        Math.Abs(region.Quad.Right - gesture.Target.Right) < 0.01 && Math.Abs(region.Quad.Bottom - gesture.Target.Bottom) < 0.01).ToArray();
                if (matches.Length == 1) {
                    var region = matches[0];
                    ApplyObjectSelection(new(PdfEditorSelectionKind.Image, gesture.Selection.PageNumber, gesture.Target,
                        ObjectNumber: region.ObjectNumber, Subtype: region.Subtype, ImagePlacement: region.ImagePlacement));
                }
            }
        } catch (Exception exception) { ErrorMessage = Infrastructure.StudioMessages.Describe(exception); }
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

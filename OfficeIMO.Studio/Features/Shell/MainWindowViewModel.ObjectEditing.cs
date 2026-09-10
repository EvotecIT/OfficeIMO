using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    [ObservableProperty]
    private string _selectedObjectText = string.Empty;

    [ObservableProperty]
    private double _objectMoveX = 10D;

    [ObservableProperty]
    private double _objectMoveY;

    [ObservableProperty]
    private double _selectedAnnotationX;

    [ObservableProperty]
    private double _selectedAnnotationY;

    [ObservableProperty]
    private double _selectedAnnotationWidth;

    [ObservableProperty]
    private double _selectedAnnotationHeight;

    [ObservableProperty]
    private string _replaceAllFindText = string.Empty;

    [ObservableProperty]
    private string _replaceAllReplacementText = string.Empty;

    [ObservableProperty]
    private bool _replaceAllMatchCase;

    [ObservableProperty]
    private bool _replaceAllWholeWords;

    public bool CanReplaceSelectedText => HasSelectedText && CanEditPageContent;

    public bool CanReplaceSelectedImage => HasSelectedImage && CanEditPageContent;

    public bool CanResizeSelectedAnnotation => HasSelectedAnnotation && CanEditAnnotations;

    [RelayCommand]
    private async Task MoveSelectedObjectAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedObject is not PdfEditorSelection selection) return;
        double deltaX = ObjectMoveX;
        double deltaY = ObjectMoveY;
        var imageOptions = new PdfImageEditOptions { Layer = PlaceEditedImageBehindContent ? PdfImageEditLayer.BehindExistingContent : PdfImageEditLayer.AboveExistingContent };
        ClearObjectSelection();
        bool succeeded = await RunMutationAsync(
            token => selection.Kind switch {
                PdfEditorSelectionKind.Text => _workspace.MoveSelectedTextAsync(selection, deltaX, deltaY, token, CreateProgress()),
                PdfEditorSelectionKind.Image => _workspace.MoveSelectedImageAsync(selection, deltaX, deltaY, token, CreateProgress(), imageOptions),
                PdfEditorSelectionKind.Annotation when selection.ObjectNumber is int objectNumber =>
                    _workspace.MoveAnnotationAsync(objectNumber, selection.PageNumber, deltaX, deltaY, token, CreateProgress()),
                _ => throw new InvalidOperationException("The selected object cannot be moved.")
            },
            cancellationToken).ConfigureAwait(true);
        if (succeeded) OperationStatus = "Selected object moved. Save when ready.";
    }

    [RelayCommand]
    private async Task ReplaceSelectedImageAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedObject is not { Kind: PdfEditorSelectionKind.Image } selection) return;
        PdfWorkspace workspace = _workspace;
        long revision = workspace.Revision;
        byte[]? bytes;
        ErrorMessage = null;
        try {
            bytes = await _pickImage(cancellationToken).ConfigureAwait(true);
        } catch (OperationCanceledException) {
            OperationStatus = UiText("Workspace.OperationCancelled");
            return;
        } catch (Exception ex) {
            ErrorMessage = ex.Message;
            OperationStatus = UiText("Workspace.OperationFailed");
            return;
        }
        if (bytes is null) return;
        if (!ReferenceEquals(_workspace, workspace) || workspace.Revision != revision) {
            OperationStatus = "The document changed while the replacement image was being selected. Select the image again.";
            return;
        }
        ClearObjectSelection();
        var imageOptions = new PdfImageEditOptions { Layer = PlaceEditedImageBehindContent ? PdfImageEditLayer.BehindExistingContent : PdfImageEditLayer.AboveExistingContent };
        bool succeeded = await RunMutationAsync(
            token => workspace.ReplaceSelectedImageAsync(selection, bytes, token, CreateProgress(), imageOptions),
            cancellationToken).ConfigureAwait(true);
        if (succeeded) OperationStatus = "Selected image replaced. Save when ready.";
    }

    [RelayCommand]
    private async Task DeleteSelectedObjectAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedObject is not PdfEditorSelection selection) return;
        ClearObjectSelection();
        bool succeeded = await RunMutationAsync(
            token => selection.Kind switch {
                PdfEditorSelectionKind.Text => _workspace.ReplaceSelectedTextAsync(selection, string.Empty, options: null, token, CreateProgress()),
                PdfEditorSelectionKind.Image => _workspace.RemoveSelectedImageAsync(selection, token, CreateProgress()),
                PdfEditorSelectionKind.Annotation when selection.ObjectNumber is int objectNumber =>
                    _workspace.RemoveAnnotationAsync(objectNumber, token, CreateProgress()),
                _ => throw new InvalidOperationException("The selected object cannot be deleted.")
            },
            cancellationToken).ConfigureAwait(true);
        if (succeeded) OperationStatus = "Selected object removed. Save when ready.";
    }

    [RelayCommand]
    private async Task ResizeSelectedAnnotationAsync(CancellationToken cancellationToken) {
        if (_workspace is null ||
            SelectedObject is not { Kind: PdfEditorSelectionKind.Annotation, ObjectNumber: int objectNumber } selection) return;
        var rectangle = new PdfPageRectangle(
            SelectedAnnotationX,
            SelectedAnnotationY,
            SelectedAnnotationX + SelectedAnnotationWidth,
            SelectedAnnotationY + SelectedAnnotationHeight);
        ClearObjectSelection();
        bool succeeded = await RunMutationAsync(
            token => _workspace.ResizeAnnotationAsync(objectNumber, selection.PageNumber, rectangle, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
        if (succeeded) OperationStatus = "Annotation geometry updated. Save when ready.";
    }
}

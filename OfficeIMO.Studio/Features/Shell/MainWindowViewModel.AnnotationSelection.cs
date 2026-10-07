using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasSelectedAnnotation))]
    [NotifyPropertyChangedFor(nameof(SelectedAnnotationObjectNumber))]
    [NotifyPropertyChangedFor(nameof(HasSelectedAnnotations))]
    [NotifyPropertyChangedFor(nameof(CanEditSelectedAnnotations))]
    [NotifyPropertyChangedFor(nameof(CanGroupAnnotations))]
    [NotifyPropertyChangedFor(nameof(CanUngroupAnnotations))]
    [NotifyPropertyChangedFor(nameof(CanResizeSelectedAnnotation))]
    private IReadOnlyList<PdfEditorSelection> _selectedAnnotations = Array.Empty<PdfEditorSelection>();

    public bool HasSelectedAnnotations => SelectedAnnotations.Count > 0;
    public bool CanEditSelectedAnnotations => HasSelectedAnnotations && CanEditAnnotations && !IsWorkspaceBusy;
    public bool CanGroupAnnotations => CanEditSelectedAnnotations && SelectedAnnotations.Count > 1 &&
        SelectedAnnotations.All(selection => selection.Subtype is not ("Link" or "Redact")) &&
        SelectedAnnotations.All(selection => _workspace?.DocumentInfo?.Annotations.FirstOrDefault(annotation =>
            annotation.ObjectNumber == selection.ObjectNumber)?.Review is not { IsReply: true });
    public bool CanUngroupAnnotations => CanEditSelectedAnnotations && SelectedAnnotations.Any(selection =>
        _workspace?.DocumentInfo?.Annotations.FirstOrDefault(annotation => annotation.ObjectNumber == selection.ObjectNumber)?.Review?.IsGroup == true);

    partial void OnSelectedAnnotationsChanged(IReadOnlyList<PdfEditorSelection> value) {
        foreach (var page in Pages) page.SelectedAnnotations = value.Where(selection => selection.PageNumber == page.PageNumber).ToArray();
    }

    private void OnPageAnnotationsSelected(PdfAnnotationSelectionRequest request) {
        if (_workspace?.DocumentInfo is not { } info || IsWorkspaceBusy) return;
        if (request.Selections.Count == 0) { if (!request.Additive) ClearObjectSelection(); return; }
        int pageNumber = request.Selections[0].PageNumber;
        if (request.Selections.Any(selection => selection.PageNumber != pageNumber || selection.ObjectNumber is null)) return;
        try {
            var logicalPage = _workspace.CreateDocumentSnapshot().Read(new PdfReadOptions { Profile = PdfReadProfile.Fast }).Pages[pageNumber - 1];
            var requested = request.Selections.SelectMany(selection => PdfAnnotationGrouping.GetMembers(info.Annotations, selection.ObjectNumber!.Value))
                .Distinct().Select(annotation => {
                    var quad = logicalPage.MapUserSpaceRectangleToVisual(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2);
                    return new PdfEditorSelection(PdfEditorSelectionKind.Annotation, pageNumber,
                        new(quad.Left, quad.Top, quad.Right, quad.Bottom), ObjectNumber: annotation.ObjectNumber, Subtype: annotation.Subtype);
                }).ToArray();
            var current = request.Additive ? SelectedAnnotations.Where(selection => selection.PageNumber == pageNumber).ToList() : [];
            bool remove = request.Additive && request.Toggle && requested.All(selection => current.Any(item => item.ObjectNumber == selection.ObjectNumber));
            foreach (var selection in requested) {
                current.RemoveAll(item => item.ObjectNumber == selection.ObjectNumber);
                if (!remove) current.Add(selection);
            }
            SelectedAnnotations = current.ToArray();
            if (current.Count == 0) { ClearObjectSelection(); return; }
            var bounds = new PdfEditorVisualBounds(current.Min(item => item.Bounds.Left), current.Min(item => item.Bounds.Top),
                current.Max(item => item.Bounds.Right), current.Max(item => item.Bounds.Bottom));
            ApplyObjectSelection(current[0] with { Bounds = bounds });
            if (current.Count > 1) SelectedObjectSummary = UiFormat("AnnotationSelection.Count", current.Count, pageNumber);
        } catch (Exception error) { ErrorMessage = error.Message; }
    }

    private async Task RunSelectedAnnotationEditAsync(Func<PdfDocumentAnnotations, IReadOnlyList<int>, PdfAnnotationEditResult> edit,
        string description, CancellationToken token) {
        if (_workspace is not { } workspace || !CanEditSelectedAnnotations) return;
        var selections = SelectedAnnotations.ToArray();
        long revision = workspace.Revision;
        ClearObjectSelection();
        await RunMutationAsync(cancellation => workspace.EditAnnotationsAsync(selections, revision, edit, description,
            cancellation, CreateProgress()), token).ConfigureAwait(true);
    }

    private async void OnAnnotationKeyRequested(Avalonia.Input.Key key, Avalonia.Input.KeyModifiers modifiers) {
        if (!CanEditSelectedAnnotations) return;
        switch (key) {
            case Avalonia.Input.Key.Delete or Avalonia.Input.Key.Back: await DeleteSelectedObjectAsync(CancellationToken.None); break;
            case Avalonia.Input.Key.D: await CopyAnnotationsAsync(CancellationToken.None); break;
            case Avalonia.Input.Key.G:
                if (modifiers.HasFlag(Avalonia.Input.KeyModifiers.Shift)) await UngroupAnnotationsAsync(CancellationToken.None);
                else await GroupAnnotationsAsync(CancellationToken.None);
                break;
            case Avalonia.Input.Key.OemCloseBrackets: await RaiseAnnotationsAsync(CancellationToken.None); break;
            case Avalonia.Input.Key.OemOpenBrackets: await LowerAnnotationsAsync(CancellationToken.None); break;
        }
    }

    [RelayCommand] private Task GroupAnnotationsAsync(CancellationToken token) => !CanGroupAnnotations ? Task.CompletedTask :
        RunSelectedAnnotationEditAsync((editor, numbers) => editor.Group(numbers), UiText("AnnotationSelection.Grouped"), token);
    [RelayCommand] private Task UngroupAnnotationsAsync(CancellationToken token) => !CanUngroupAnnotations ? Task.CompletedTask :
        RunSelectedAnnotationEditAsync((editor, numbers) => editor.Ungroup(numbers), UiText("AnnotationSelection.Ungrouped"), token);
    [RelayCommand] private Task CopyAnnotationsAsync(CancellationToken token) =>
        RunSelectedAnnotationEditAsync((editor, numbers) => editor.CopyMany(numbers), UiText("AnnotationSelection.Copied"), token);
    [RelayCommand] private Task RaiseAnnotationsAsync(CancellationToken token) =>
        RunSelectedAnnotationEditAsync((editor, numbers) => editor.Arrange(numbers, PdfAnnotationOrderChange.Raise), UiText("AnnotationSelection.Raised"), token);
    [RelayCommand] private Task LowerAnnotationsAsync(CancellationToken token) =>
        RunSelectedAnnotationEditAsync((editor, numbers) => editor.Arrange(numbers, PdfAnnotationOrderChange.Lower), UiText("AnnotationSelection.Lowered"), token);
}

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
    [NotifyPropertyChangedFor(nameof(HasCrossPageAnnotationSelection))]
    private IReadOnlyList<PdfEditorSelection> _selectedAnnotations = Array.Empty<PdfEditorSelection>();

    public bool HasSelectedAnnotations => SelectedAnnotations.Count > 0;
    public bool HasCrossPageAnnotationSelection => SelectedAnnotations.Select(selection => selection.PageNumber).Distinct().Skip(1).Any();
    public bool CanEditSelectedAnnotations => HasSelectedAnnotations && CanEditAnnotations && !IsWorkspaceBusy;
    public bool CanGroupAnnotations => CanEditSelectedAnnotations && SelectedAnnotations.Count > 1 && !HasCrossPageAnnotationSelection &&
        SelectedAnnotations.All(selection => selection.Subtype is not ("Link" or "Redact")) &&
        SelectedAnnotations.All(selection => _workspace?.AnnotationMetadata.FirstOrDefault(annotation =>
            annotation.ObjectNumber == selection.ObjectNumber)?.Review is not { IsReply: true });
    public bool CanUngroupAnnotations => CanEditSelectedAnnotations && SelectedAnnotations.Any(selection =>
        _workspace?.AnnotationMetadata.FirstOrDefault(annotation => annotation.ObjectNumber == selection.ObjectNumber)?.Review?.IsGroup == true);

    partial void OnSelectedAnnotationsChanged(IReadOnlyList<PdfEditorSelection> value) {
        foreach (var page in Pages) page.SelectedAnnotations = value;
    }

    internal void OnPageAnnotationsSelected(PdfAnnotationSelectionRequest request) {
        if (_workspace is not { } workspace || IsWorkspaceBusy) return;
        if (request.Selections.Count == 0) { if (!request.Additive) ClearObjectSelection(); return; }
        if (request.Selections.Any(selection => selection.ObjectNumber is null || selection.PageNumber < 1 || selection.PageNumber > Pages.Count)) return;
        try {
            var layouts = _workspace.CreateDocumentSnapshot().GetPageLayouts();
            var requested = request.Selections.SelectMany(selection => PdfAnnotationGrouping.GetMembers(workspace.AnnotationMetadata, selection.ObjectNumber!.Value))
                .Distinct().Select(annotation => {
                    int pageNumber = annotation.PageNumber!.Value;
                    var quad = layouts[pageNumber - 1].MapUserSpaceRectangleToVisual(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2);
                    return new PdfEditorSelection(PdfEditorSelectionKind.Annotation, pageNumber,
                        new(quad.Left, quad.Top, quad.Right, quad.Bottom), ObjectNumber: annotation.ObjectNumber, Subtype: annotation.Subtype);
                }).ToArray();
            var current = request.Additive ? SelectedAnnotations.ToList() : [];
            bool remove = request.Additive && request.Toggle && requested.All(selection => current.Any(item => item.ObjectNumber == selection.ObjectNumber));
            foreach (var selection in requested) {
                current.RemoveAll(item => item.ObjectNumber == selection.ObjectNumber);
                if (!remove) current.Add(selection);
            }
            SelectedAnnotations = current.ToArray();
            if (current.Count == 0) { ClearObjectSelection(); return; }
            ApplyObjectSelection(BoundPageSelection(current.Where(item => item.PageNumber == current[0].PageNumber).ToArray()));
            // Bounds are page-local. Never merge rectangles from different coordinate systems.
            foreach (var page in Pages) {
                var members = current.Where(item => item.PageNumber == page.PageNumber).ToArray();
                page.SelectedObject = members.Length == 0 ? null : BoundPageSelection(members);
            }
            if (current.Count > 1) SelectedObjectSummary = HasCrossPageAnnotationSelection
                ? UiFormat("AnnotationSelection.CrossPageCount", current.Count, current.Select(item => item.PageNumber).Distinct().Count())
                : UiFormat("AnnotationSelection.Count", current.Count, current[0].PageNumber);
        } catch (Exception error) { ErrorMessage = error.Message; }
    }

    private static PdfEditorSelection BoundPageSelection(IReadOnlyList<PdfEditorSelection> members) => members[0] with {
        Bounds = new(members.Min(item => item.Bounds.Left), members.Min(item => item.Bounds.Top),
            members.Max(item => item.Bounds.Right), members.Max(item => item.Bounds.Bottom))
    };

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
        RunSelectedAnnotationEditAsync((editor, numbers) => editor.CopyManyVisual(numbers), UiText("AnnotationSelection.Copied"), token);
    [RelayCommand] private Task RaiseAnnotationsAsync(CancellationToken token) =>
        RunSelectedAnnotationEditAsync((editor, numbers) => editor.Arrange(numbers, PdfAnnotationOrderChange.Raise), UiText("AnnotationSelection.Raised"), token);
    [RelayCommand] private Task LowerAnnotationsAsync(CancellationToken token) =>
        RunSelectedAnnotationEditAsync((editor, numbers) => editor.Arrange(numbers, PdfAnnotationOrderChange.Lower), UiText("AnnotationSelection.Lowered"), token);
}

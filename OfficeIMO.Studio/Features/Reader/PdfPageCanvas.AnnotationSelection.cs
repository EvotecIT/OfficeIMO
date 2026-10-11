using Avalonia;
using Avalonia.Input;
using Avalonia.Media;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageCanvas {
    public static readonly StyledProperty<IReadOnlyList<PdfEditorSelection>> SelectedAnnotationsProperty =
        AvaloniaProperty.Register<PdfPageCanvas, IReadOnlyList<PdfEditorSelection>>(nameof(SelectedAnnotations), Array.Empty<PdfEditorSelection>());

    /// <summary>Document annotation selection; drawing uses only this page's members. Page-organizer selection is independent.</summary>
    public IReadOnlyList<PdfEditorSelection> SelectedAnnotations {
        get => GetValue(SelectedAnnotationsProperty);
        set => SetValue(SelectedAnnotationsProperty, value);
    }

    internal event Action<PdfAnnotationSelectionRequest>? AnnotationSelectionRequested;
    internal event Action<Key, KeyModifiers>? AnnotationKeyRequested;
    private bool _additiveAnnotationSelection;

    private static bool IsAdditive(KeyModifiers modifiers) => (modifiers & (KeyModifiers.Shift | KeyModifiers.Control | KeyModifiers.Meta)) != 0;

    private void SelectAnnotationRegion(PdfPageInteractionRegion region, bool additive, bool toggle = true) {
        if (Scene is not null) AnnotationSelectionRequested?.Invoke(new([CreateSelection(Scene.PageNumber, region)], additive, toggle));
    }

    private void SelectAnnotationsInRectangle() {
        if (Scene is null || !_selectionStart.HasValue || !_selectionEnd.HasValue) return;
        var request = RaiseAnnotationMarquee(completed: true);
        if (!request.Handled) AnnotationSelectionRequested?.Invoke(new(
            GetAnnotationsInControlRectangle(new Rect(_selectionStart.Value, _selectionEnd.Value).Normalize()), _additiveAnnotationSelection, Toggle: false));
        _selectionStart = null; _selectionEnd = null;
    }

    private void SelectAllAnnotations() {
        if (Scene?.Interactions is not { } interactions) return;
        AnnotationSelectionRequested?.Invoke(new(interactions.Regions.Where(region => region.Kind == PdfInteractionKind.Annotation && region.ObjectNumber.HasValue)
            .Select(region => CreateSelection(Scene.PageNumber, region)).ToArray(), false, Toggle: false));
    }

    private bool HandleAnnotationKey(KeyEventArgs e) {
        if (SelectionMode != PdfEditorSelectionMode.Annotations) return false;
        bool command = (e.KeyModifiers & (KeyModifiers.Control | KeyModifiers.Meta)) != 0;
        if (command && e.Key == Key.A) SelectAllAnnotations();
        else if (SelectedAnnotations.Count > 0 && (e.Key == Key.Delete || e.Key == Key.Back || command && e.Key is Key.G or Key.D ||
            command && e.Key is Key.OemOpenBrackets or Key.OemCloseBrackets)) AnnotationKeyRequested?.Invoke(e.Key, e.KeyModifiers);
        else if (SelectedAnnotations.Count > 0 && command && e.Key is Key.Left or Key.Right or Key.Up or Key.Down && SelectedObject is { } selected) {
            double amount = e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? 10D : 1D;
            var delta = e.Key switch { Key.Left => new Vector(-amount, 0), Key.Right => new Vector(amount, 0), Key.Up => new Vector(0, -amount), _ => new Vector(0, amount) };
            var bounds = selected.Bounds;
            ObjectTransformCompleted?.Invoke(new(selected, new(bounds.Left + delta.X, bounds.Top + delta.Y, bounds.Right + delta.X, bounds.Bottom + delta.Y), SelectedAnnotations.ToArray()));
        } else if (e.Key is Key.Enter or Key.Space && IsAdditive(e.KeyModifiers)) {
            var interactions = GetKeyboardInteractions();
            if (_keyboardInteractionIndex < 0 || _keyboardInteractionIndex >= interactions.Count) return false;
            SelectAnnotationRegion(interactions[_keyboardInteractionIndex], true);
        } else return false;
        return true;
    }

    private void DrawAnnotationSelections(DrawingContext context) {
        if (SelectedAnnotations.Count <= 1) return;
        var pen = new Pen(new SolidColorBrush(PageAccent), 1D);
        foreach (var selection in SelectedAnnotations.Where(item => item.PageNumber == Scene?.PageNumber)) {
            var bounds = selection.Bounds;
            context.DrawRectangle(null, pen, new Rect(bounds.Left, bounds.Top, bounds.Width, bounds.Height));
        }
    }
}

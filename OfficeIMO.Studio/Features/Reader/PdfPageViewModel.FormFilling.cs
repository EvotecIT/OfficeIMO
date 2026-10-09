using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>Places shared field drafts over their visible widgets using the PDF engine's page geometry.</summary>
public sealed partial class PdfPageViewModel {
    private PdfPageInteractionMap? _inlineFormMap;
    private IReadOnlyList<PdfFormFieldViewModel> _inlineFormFields = [];
    [ObservableProperty] private PdfFormFieldViewModel? _inlineFormField;
    [ObservableProperty] private int? _formAnchorObjectNumber;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasInlineFormEditor))]
    private IReadOnlyList<PdfInlineFormWidgetViewModel> _inlineFormWidgets = [];

    public bool HasInlineFormEditor => InlineFormWidgets.Count > 0;
    /// <summary>Set when a page click or Tab should focus the selected widget after layout.</summary>
    internal bool FocusInlineFormEditorRequested { get; set; }
    /// <summary>Acknowledges successful presentation focus to the canonical page, so it cannot replay the request.</summary>
    internal event Action? InlineFormEditorFocusCompleted;

    internal void CompleteInlineFormEditorFocus() {
        FocusInlineFormEditorRequested = false;
        InlineFormEditorFocusCompleted?.Invoke();
    }
    /// <summary>Raised by Tab (+1) and Shift+Tab (-1) inside an on-page editor.</summary>
    internal event Action<int>? InlineFormNavigationRequested;
    internal void RequestInlineFormNavigation(int direction) => InlineFormNavigationRequested?.Invoke(direction);

    internal void ShowInlineFormFields(IReadOnlyList<PdfFormFieldViewModel> fields, PdfFormFieldViewModel? field, bool focus) {
        if (field is null) FocusInlineFormEditorRequested = false;
        else if (focus) FocusInlineFormEditorRequested = true;
        if (!ReferenceEquals(InlineFormField, field)) {
            InlineFormField = field;
            FormAnchorObjectNumber = null;
        }
        if (!_inlineFormFields.SequenceEqual(fields)) {
            _inlineFormFields = fields;
            _inlineFormMap = null;
        }
        UpdateInlineFormBounds();
        if (field is not null && focus) OnPropertyChanged(nameof(FocusInlineFormEditorRequested));
    }

    private void UpdateInlineFormBounds() {
        if (_inlineFormFields.Count == 0 || Scene is not { Interactions: { } map } scene) {
            InlineFormWidgets = [];
            _inlineFormMap = null;
            return;
        }
        if (!ReferenceEquals(_inlineFormMap, map)) {
            var fields = _inlineFormFields.ToDictionary(field => field.Name, StringComparer.Ordinal);
            InlineFormWidgets = map.Regions.Where(region => region.Kind == PdfInteractionKind.FormWidget)
                .Select(region => region.FieldName is not null && fields.TryGetValue(region.FieldName, out var field)
                    ? PdfInlineFormWidgetViewModel.Create(this, field, region) : null)
                .OfType<PdfInlineFormWidgetViewModel>().ToArray();
            _inlineFormMap = map;
        }
        foreach (var widget in InlineFormWidgets) widget.UpdateBounds(CanvasWidth / Math.Max(1D, scene.Drawing.Width), CanvasHeight / Math.Max(1D, scene.Drawing.Height));
    }
}

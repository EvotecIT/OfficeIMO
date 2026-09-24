using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>Fills the selected form field directly on the page, over the field's own widget.</summary>
public sealed partial class PdfPageViewModel {
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasInlineFormEditor))]
    private PdfFormFieldViewModel? _inlineFormField;

    [ObservableProperty] private double _inlineFormLeft;
    [ObservableProperty] private double _inlineFormTop;
    [ObservableProperty] private double _inlineFormWidth;
    [ObservableProperty] private double _inlineFormHeight;
    [ObservableProperty] private double _inlineFormFontSize = 12D;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasInlineFormEditor))]
    private bool _hasInlineFormBounds;

    public bool HasInlineFormEditor => InlineFormField is not null && HasInlineFormBounds;

    /// <summary>Set when the editor should take keyboard focus as soon as it appears (page click or Tab).</summary>
    internal bool FocusInlineFormEditorRequested { get; set; }

    /// <summary>Raised by Tab (+1) and Shift+Tab (-1) inside the on-page editor.</summary>
    internal event Action<int>? InlineFormNavigationRequested;

    internal void RequestInlineFormNavigation(int direction) => InlineFormNavigationRequested?.Invoke(direction);

    internal void ShowInlineFormField(PdfFormFieldViewModel? field, bool focus) {
        // A pending focus request survives repeated refreshes until the view consumes it.
        if (field is null) FocusInlineFormEditorRequested = false;
        else if (focus) FocusInlineFormEditorRequested = true;
        if (!ReferenceEquals(InlineFormField, field)) InlineFormField = field;
        UpdateInlineFormBounds();
        if (field is not null && focus) OnPropertyChanged(nameof(FocusInlineFormEditorRequested));
    }

    private void UpdateInlineFormBounds() {
        if (InlineFormField is not { } field || Scene is not { Interactions: { } map } scene) {
            HasInlineFormBounds = false;
            return;
        }
        PdfPageInteractionRegion? region = map.Regions.FirstOrDefault(candidate =>
            candidate.Kind == PdfInteractionKind.FormWidget && candidate.FieldName == field.Name);
        if (region is null) {
            HasInlineFormBounds = false;
            return;
        }
        double scaleX = DisplayWidth / Math.Max(1D, scene.Drawing.Width);
        double scaleY = DisplayHeight / Math.Max(1D, scene.Drawing.Height);
        InlineFormLeft = region.Quad.Left * scaleX;
        InlineFormTop = region.Quad.Top * scaleY;
        InlineFormWidth = Math.Max(field.IsCheckBoxEditor ? 18D : 60D, region.Quad.Width * scaleX);
        InlineFormHeight = Math.Max(field.IsCheckBoxEditor ? 18D : 22D, region.Quad.Height * scaleY);
        InlineFormFontSize = Math.Clamp(InlineFormHeight * (field.IsMultiline ? 0.3D : 0.55D), 9D, 16D);
        HasInlineFormBounds = true;
    }
}

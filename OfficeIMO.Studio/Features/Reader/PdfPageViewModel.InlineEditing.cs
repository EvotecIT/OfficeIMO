using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageViewModel {
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasInlineTextDraft))]
    private PdfTextDraftViewModel? _inlineTextDraft;
    public bool HasInlineTextDraft => InlineTextDraft is not null;
    public double InlineEditorWidth => Math.Min(Math.Max(240, (SelectedObject?.Bounds.Width ?? 240) * _zoom), Math.Max(1, DisplayWidth - 8));
    public double InlineEditorLeft => Math.Clamp((SelectedObject?.Bounds.Left ?? 0) * _zoom, 0, Math.Max(0, DisplayWidth - InlineEditorWidth - 4));
    public double InlineEditorTop => Math.Clamp((SelectedObject?.Bounds.Top ?? 0) * _zoom, 0, Math.Max(0, DisplayHeight - 160));
    partial void OnSelectedObjectChanged(PdfEditorSelection? value) => UpdateInlineEditorPosition();
    partial void OnDisplayWidthChanged(double value) => UpdateInlineEditorPosition();
    partial void OnDisplayHeightChanged(double value) => UpdateInlineEditorPosition();
    private void UpdateInlineEditorPosition() {
        OnPropertyChanged(nameof(InlineEditorWidth));
        OnPropertyChanged(nameof(InlineEditorLeft));
        OnPropertyChanged(nameof(InlineEditorTop));
    }
}

using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageViewModel {
    private double? _canvasWidth;
    private double? _canvasHeight;
    private double _visualWidth;
    private double _visualHeight;
    internal double VisualPageWidth => Scene?.Drawing.Width ?? _visualWidth;
    internal double VisualPageHeight => Scene?.Drawing.Height ?? _visualHeight;
    private double CanvasWidth => Math.Max(1, _canvasWidth ?? DisplayWidth - 2);
    private double CanvasHeight => Math.Max(1, _canvasHeight ?? DisplayHeight - 2);
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasInlineTextDraft))]
    private PdfTextDraftViewModel? _inlineTextDraft;
    public bool HasInlineTextDraft => InlineTextDraft is not null;
    public double InlineEditorWidth => Math.Min(Math.Max(240, (SelectedObject?.Bounds.Width ?? 240) * CanvasWidth / VisualPageWidth), Math.Max(1, CanvasWidth - 8));
    public double InlineEditorLeft => Math.Clamp((SelectedObject?.Bounds.Left ?? 0) * CanvasWidth / VisualPageWidth, 0, Math.Max(0, CanvasWidth - InlineEditorWidth - 4));
    public double InlineEditorTop => Math.Clamp((SelectedObject?.Bounds.Top ?? 0) * CanvasHeight / VisualPageHeight, 0, Math.Max(0, CanvasHeight - 160));
    internal void UpdateCanvasSize(Avalonia.Size size) {
        _canvasWidth = size.Width > 0 ? size.Width : null;
        _canvasHeight = size.Height > 0 ? size.Height : null;
        UpdateInlineEditorPosition();
    }
    partial void OnSceneChanged(PdfPageScene? value) => UpdateInlineEditorPosition();
    partial void OnSelectedObjectChanged(PdfEditorSelection? value) => UpdateInlineEditorPosition();
    partial void OnDisplayWidthChanged(double value) => UpdateInlineEditorPosition();
    partial void OnDisplayHeightChanged(double value) => UpdateInlineEditorPosition();
    private void UpdateInlineEditorPosition() {
        OnPropertyChanged(nameof(InlineEditorWidth));
        OnPropertyChanged(nameof(InlineEditorLeft));
        OnPropertyChanged(nameof(InlineEditorTop));
    }
}

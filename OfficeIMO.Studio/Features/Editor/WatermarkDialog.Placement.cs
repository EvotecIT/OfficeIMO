using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;

namespace OfficeIMO.Studio.Features.Editor;

public sealed partial class WatermarkDialog {
    private Point? _dragStart;
    private Point _placementStart;
    private Point _placementCurrent;
    private Rect _previewPageBounds;
    private double _previewScale;

    private void OnPreviewPressed(object? sender, PointerPressedEventArgs e) {
        if (DataContext is not WatermarkPreviewViewModel { CanApply: true, Prepared: { } preview } model
            || !e.GetCurrentPoint(PreviewSurface).Properties.IsLeftButtonPressed) return;
        _previewScale = Math.Min(PreviewSurface.Bounds.Width / preview.PageWidth,
            PreviewSurface.Bounds.Height / preview.PageHeight);
        if (!double.IsFinite(_previewScale) || _previewScale <= 0) return;
        double width = preview.PageWidth * _previewScale;
        double height = preview.PageHeight * _previewScale;
        _previewPageBounds = new Rect((PreviewSurface.Bounds.Width - width) / 2,
            (PreviewSurface.Bounds.Height - height) / 2, width, height);
        Point point = e.GetPosition(PreviewSurface);
        if (!_previewPageBounds.Contains(point)) return;
        _dragStart = point;
        _placementStart = new Point((double?)model.X ?? (preview.PageWidth - (double)model.Width) / 2,
            (double?)model.Y ?? (preview.PageHeight - (double)model.Height) / 2);
        _placementCurrent = _placementStart;
        ShowPlacement(model);
        e.Pointer.Capture(PreviewSurface);
        e.Handled = true;
    }

    private void OnPreviewMoved(object? sender, PointerEventArgs e) {
        if (_dragStart is not { } start || DataContext is not WatermarkPreviewViewModel model) return;
        Point point = e.GetPosition(PreviewSurface);
        _placementCurrent = new Point(
            Math.Clamp(_placementStart.X + (point.X - start.X) / _previewScale, 0,
                Math.Max(0, _previewPageBounds.Width / _previewScale - (double)model.Width)),
            Math.Clamp(_placementStart.Y + (point.Y - start.Y) / _previewScale, 0,
                Math.Max(0, _previewPageBounds.Height / _previewScale - (double)model.Height)));
        ShowPlacement(model);
        e.Handled = true;
    }

    private async void OnPreviewReleased(object? sender, PointerReleasedEventArgs e) {
        if (_dragStart is null || DataContext is not WatermarkPreviewViewModel model) return;
        Point placement = _placementCurrent;
        bool moved = Math.Abs(_placementCurrent.X - _placementStart.X) > 0.5
            || Math.Abs(_placementCurrent.Y - _placementStart.Y) > 0.5;
        _dragStart = null;
        PlacementOutline.IsVisible = false;
        e.Pointer.Capture(null);
        e.Handled = true;
        if (!moved) return;
        model.X = (decimal)Math.Round(placement.X, 2);
        model.Y = (decimal)Math.Round(placement.Y, 2);
        await model.PreviewCommand.ExecuteAsync(null);
    }

    private void OnPreviewCaptureLost(object? sender, PointerCaptureLostEventArgs e) {
        _dragStart = null;
        PlacementOutline.IsVisible = false;
    }

    private void ShowPlacement(WatermarkPreviewViewModel model) {
        PlacementOutline.Width = (double)model.Width * _previewScale;
        PlacementOutline.Height = (double)model.Height * _previewScale;
        Canvas.SetLeft(PlacementOutline, _previewPageBounds.X + _placementCurrent.X * _previewScale);
        Canvas.SetTop(PlacementOutline, _previewPageBounds.Y + _placementCurrent.Y * _previewScale);
        PlacementOutline.RenderTransform = new RotateTransform((double)model.Rotation);
        PlacementOutline.IsVisible = true;
    }
}

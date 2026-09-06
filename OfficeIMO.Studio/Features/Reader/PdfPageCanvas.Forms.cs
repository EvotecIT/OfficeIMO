using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageCanvas {
    public static readonly StyledProperty<string?> FormAnchorFieldNameProperty =
        AvaloniaProperty.Register<PdfPageCanvas, string?>(nameof(FormAnchorFieldName));

    /// <summary>Named field to highlight using the engine's visual widget geometry.</summary>
    public string? FormAnchorFieldName {
        get => GetValue(FormAnchorFieldNameProperty);
        set => SetValue(FormAnchorFieldNameProperty, value);
    }

    internal IReadOnlyList<Rect> FormAnchorBounds => string.IsNullOrEmpty(FormAnchorFieldName) ? [] :
        Scene?.Interactions.Regions.Where(region => region.Kind == PdfInteractionKind.FormWidget && region.FieldName == FormAnchorFieldName)
            .Select(region => new Rect(region.Quad.Left, region.Quad.Top, region.Quad.Width, region.Quad.Height)).ToArray() ?? [];

    private void DrawFormAnchor(DrawingContext context) {
        foreach (Rect bounds in FormAnchorBounds) context.DrawRectangle(null, new Pen(Brushes.DodgerBlue, 2D), bounds.Inflate(3D));
    }

    private void QueueFormAnchorReveal() => Dispatcher.UIThread.Post(() => {
        if (_disposed || Scene is not { } scene || FormAnchorBounds.FirstOrDefault() is not { Width: > 0, Height: > 0 } area ||
            Bounds.Width <= 0 || Bounds.Height <= 0) return;
        double x = Bounds.Width / Math.Max(1D, scene.Drawing.Width);
        double y = Bounds.Height / Math.Max(1D, scene.Drawing.Height);
        this.BringIntoView(new Rect(area.X * x, area.Y * y, area.Width * x, area.Height * y).Inflate(20D));
    }, DispatcherPriority.Loaded);
}

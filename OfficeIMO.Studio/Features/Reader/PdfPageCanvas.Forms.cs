using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageCanvas {
    public static readonly StyledProperty<string?> FormAnchorFieldNameProperty =
        AvaloniaProperty.Register<PdfPageCanvas, string?>(nameof(FormAnchorFieldName));

    public static readonly StyledProperty<int?> FormAnchorObjectNumberProperty =
        AvaloniaProperty.Register<PdfPageCanvas, int?>(nameof(FormAnchorObjectNumber));

    /// <summary>The widget to reveal when the selected field has more than one appearance.</summary>
    public int? FormAnchorObjectNumber {
        get => GetValue(FormAnchorObjectNumberProperty);
        set => SetValue(FormAnchorObjectNumberProperty, value);
    }

    /// <summary>Named field to highlight using the engine's visual widget geometry.</summary>
    public string? FormAnchorFieldName {
        get => GetValue(FormAnchorFieldNameProperty);
        set => SetValue(FormAnchorFieldNameProperty, value);
    }

    internal IReadOnlyList<Rect> FormAnchorBounds => string.IsNullOrEmpty(FormAnchorFieldName) ? [] :
        Scene?.Interactions?.Regions.Where(region => region.Kind == PdfInteractionKind.FormWidget && region.FieldName == FormAnchorFieldName)
            .Select(region => new Rect(region.Quad.Left, region.Quad.Top, region.Quad.Width, region.Quad.Height)).ToArray() ?? [];

    private void DrawFormAnchor(DrawingContext context) {
        foreach (Rect bounds in FormAnchorBounds) context.DrawRectangle(null, new Pen(new SolidColorBrush(PageAccent), 2D), bounds.Inflate(3D));
    }

    private void QueueFormAnchorReveal() => this.Dispatcher.Post(() => {
        if (_disposed || Scene is not { } scene || Bounds.Width <= 0 || Bounds.Height <= 0) return;
        var region = scene.Interactions?.Regions.FirstOrDefault(region => region.Kind == PdfInteractionKind.FormWidget &&
            region.FieldName == FormAnchorFieldName && region.ObjectNumber == FormAnchorObjectNumber);
        Rect area = region is null ? FormAnchorBounds.FirstOrDefault() : new(region.Quad.Left, region.Quad.Top, region.Quad.Width, region.Quad.Height);
        if (area.Width <= 0 || area.Height <= 0) return;
        double x = Bounds.Width / Math.Max(1D, scene.Drawing.Width);
        double y = Bounds.Height / Math.Max(1D, scene.Drawing.Height);
        this.BringIntoView(new Rect(area.X * x, area.Y * y, area.Width * x, area.Height * y).Inflate(20D));
    }, DispatcherPriority.Loaded);
}

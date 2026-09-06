using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.Immutable;
using Avalonia.Threading;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageCanvas {
    public static readonly StyledProperty<IReadOnlyList<Rect>> SearchHighlightsProperty =
        AvaloniaProperty.Register<PdfPageCanvas, IReadOnlyList<Rect>>(nameof(SearchHighlights), Array.Empty<Rect>());

    public static readonly StyledProperty<Rect?> ActiveSearchHighlightProperty =
        AvaloniaProperty.Register<PdfPageCanvas, Rect?>(nameof(ActiveSearchHighlight));

    public IReadOnlyList<Rect> SearchHighlights {
        get => GetValue(SearchHighlightsProperty);
        set => SetValue(SearchHighlightsProperty, value);
    }

    public Rect? ActiveSearchHighlight {
        get => GetValue(ActiveSearchHighlightProperty);
        set => SetValue(ActiveSearchHighlightProperty, value);
    }

    private static readonly IBrush SearchBrush = new ImmutableSolidColorBrush(Color.FromArgb(85, 255, 205, 35));
    private static readonly IBrush ActiveSearchBrush = new ImmutableSolidColorBrush(Color.FromArgb(110, 255, 145, 0));

    private void DrawSearchHighlights(DrawingContext context) {
        foreach (Rect area in SearchHighlights) context.DrawRectangle(SearchBrush, null, area);
        if (ActiveSearchHighlight is Rect selected) context.DrawRectangle(ActiveSearchBrush, new Pen(Brushes.DarkOrange, 1.5D), selected);
    }

    private void QueueSearchReveal() => Dispatcher.UIThread.Post(() => {
        if (_disposed || Scene is not { } scene || ActiveSearchHighlight is not Rect area || Bounds.Width <= 0 || Bounds.Height <= 0) return;
        double x = Bounds.Width / Math.Max(1D, scene.Drawing.Width);
        double y = Bounds.Height / Math.Max(1D, scene.Drawing.Height);
        this.BringIntoView(new Rect(area.X * x, area.Y * y, area.Width * x, area.Height * y).Inflate(20D));
    }, DispatcherPriority.Loaded);
}

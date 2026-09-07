using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageCanvas {
    public static readonly StyledProperty<int?> CommentAnchorObjectNumberProperty =
        AvaloniaProperty.Register<PdfPageCanvas, int?>(nameof(CommentAnchorObjectNumber));

    /// <summary>Review-thread root to reveal using the engine's rotated and cropped visual coordinates.</summary>
    public int? CommentAnchorObjectNumber {
        get => GetValue(CommentAnchorObjectNumberProperty);
        set => SetValue(CommentAnchorObjectNumberProperty, value);
    }

    private Rect? CommentAnchorBounds {
        get {
            if (CommentAnchorObjectNumber is not int number) return null;
            var region = Scene?.Interactions?.Regions.FirstOrDefault(region => region.Kind == PdfInteractionKind.Annotation && region.ObjectNumber == number);
            return region is null ? null : new Rect(region.Quad.Left, region.Quad.Top, region.Quad.Width, region.Quad.Height);
        }
    }

    private void DrawCommentAnchor(DrawingContext context) {
        if (CommentAnchorBounds is Rect area) context.DrawRectangle(null, new Pen(Brushes.DodgerBlue, 2D), area.Inflate(4D));
    }

    private void QueueCommentAnchorReveal() => Dispatcher.UIThread.Post(() => {
        if (_disposed || Scene is not { } scene || CommentAnchorBounds is not Rect area || Bounds.Width <= 0 || Bounds.Height <= 0) return;
        double x = Bounds.Width / Math.Max(1D, scene.Drawing.Width);
        double y = Bounds.Height / Math.Max(1D, scene.Drawing.Height);
        this.BringIntoView(new Rect(area.X * x, area.Y * y, area.Width * x, area.Height * y).Inflate(20D));
    }, DispatcherPriority.Loaded);
}

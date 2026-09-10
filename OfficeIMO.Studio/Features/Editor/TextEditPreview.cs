using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace OfficeIMO.Studio.Features.Editor;

/// <summary>Shows a magnified region of a prepared page without modifying its rendered pixels.</summary>
public sealed class TextEditPreview : Control {
    public static readonly StyledProperty<IImage?> SourceProperty = AvaloniaProperty.Register<TextEditPreview, IImage?>(nameof(Source));
    public static readonly StyledProperty<Rect?> RegionProperty = AvaloniaProperty.Register<TextEditPreview, Rect?>(nameof(Region));
    static TextEditPreview() => AffectsRender<TextEditPreview>(SourceProperty, RegionProperty);
    public IImage? Source { get => GetValue(SourceProperty); set => SetValue(SourceProperty, value); }
    /// <summary>Normalized page region, or null for the whole page.</summary>
    public Rect? Region { get => GetValue(RegionProperty); set => SetValue(RegionProperty, value); }

    public override void Render(DrawingContext context) {
        base.Render(context);
        if (Source is not { } source || Bounds.Width <= 0 || Bounds.Height <= 0) return;
        Rect region = Region ?? new Rect(0, 0, 1, 1);
        var crop = new Rect(region.X * source.Size.Width, region.Y * source.Size.Height,
            region.Width * source.Size.Width, region.Height * source.Size.Height).Intersect(new Rect(source.Size));
        if (crop.Width <= 0 || crop.Height <= 0) return;
        double scale = Math.Min(Bounds.Width / crop.Width, Bounds.Height / crop.Height);
        var destination = new Rect((Bounds.Width - crop.Width * scale) / 2, (Bounds.Height - crop.Height * scale) / 2,
            crop.Width * scale, crop.Height * scale);
        context.DrawRectangle(Brushes.White, null, destination);
        context.DrawImage(source, crop, destination);
    }
}

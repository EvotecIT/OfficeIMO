namespace OfficeIMO.DjVu;

public sealed partial class DjVuPage {
    /// <summary>Native pixel width after the declared display rotation.</summary>
    public int DisplayWidth => Rotation == 90 || Rotation == 270 ? Height : Width;
    /// <summary>Native pixel height after the declared display rotation.</summary>
    public int DisplayHeight => Rotation == 90 || Rotation == 270 ? Width : Height;
    /// <summary>Maps unrotated bottom-left source bounds to rotated top-left display bounds, in native pixels.</summary>
    public DjVuRectangle GetDisplayBounds(DjVuRectangle sourceBounds) {
        checked {
            switch (Rotation) {
                case 90: return new DjVuRectangle(sourceBounds.Y, sourceBounds.X, sourceBounds.Height, sourceBounds.Width);
                case 180: return new DjVuRectangle(Width - sourceBounds.X - sourceBounds.Width, sourceBounds.Y, sourceBounds.Width, sourceBounds.Height);
                case 270: return new DjVuRectangle(Height - sourceBounds.Y - sourceBounds.Height, Width - sourceBounds.X - sourceBounds.Width, sourceBounds.Height, sourceBounds.Width);
                default: return new DjVuRectangle(sourceBounds.X, Height - sourceBounds.Y - sourceBounds.Height, sourceBounds.Width, sourceBounds.Height);
            }
        }
    }
}

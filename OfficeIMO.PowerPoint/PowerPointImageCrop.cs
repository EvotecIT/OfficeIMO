namespace OfficeIMO.PowerPoint;

internal readonly struct PowerPointImageCrop {
    public static PowerPointImageCrop None { get; } = new(0D, 0D, 0D, 0D);

    public PowerPointImageCrop(double left, double top, double right, double bottom) {
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
    }

    public double Left { get; }

    public double Top { get; }

    public double Right { get; }

    public double Bottom { get; }

    public bool HasCrop => Left > 0D || Top > 0D || Right > 0D || Bottom > 0D;
}

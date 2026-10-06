namespace OfficeIMO.Studio.Features.Reader;

/// <summary>Visual page coordinates, with an origin at the top left, independent of a UI framework.</summary>
public readonly record struct StudioRectangle(double X, double Y, double Width, double Height) {
    /// <summary>The right edge in page coordinates.</summary>
    public double Right => X + Width;
    /// <summary>The bottom edge in page coordinates.</summary>
    public double Bottom => Y + Height;
}

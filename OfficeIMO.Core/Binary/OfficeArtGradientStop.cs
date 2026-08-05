namespace OfficeIMO.Drawing.Binary;

/// <summary>Represents one MSOSHADECOLOR entry from an OfficeArt gradient-stop array.</summary>
public sealed class OfficeArtGradientStop {
    internal OfficeArtGradientStop(OfficeArtColorReference color, double position) {
        Color = color;
        Position = position;
    }

    /// <summary>Gets the OfficeArt color reference.</summary>
    public OfficeArtColorReference Color { get; }

    /// <summary>Gets the relative position from zero through one.</summary>
    public double Position { get; }
}

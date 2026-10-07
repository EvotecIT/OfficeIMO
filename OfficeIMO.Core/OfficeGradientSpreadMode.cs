namespace OfficeIMO.Drawing;

/// <summary>Controls gradient colors beyond the authored stop interval.</summary>
public enum OfficeGradientSpreadMode {
    /// <summary>Continue the nearest endpoint color.</summary>
    Pad,
    /// <summary>Repeat the stop interval in the same direction.</summary>
    Repeat,
    /// <summary>Repeat the stop interval in alternating directions.</summary>
    Reflect
}

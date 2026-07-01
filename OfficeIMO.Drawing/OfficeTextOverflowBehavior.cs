namespace OfficeIMO.Drawing;

/// <summary>
/// Defines how measured text layout should represent content that exceeds the available bounds.
/// </summary>
public enum OfficeTextOverflowBehavior {
    /// <summary>Shorten overflowing text with an ellipsis so the measured text fits the requested bounds.</summary>
    Ellipsis,

    /// <summary>Keep overflowing text intact and rely on the caller's clipping surface to hide content outside the bounds.</summary>
    Clip
}

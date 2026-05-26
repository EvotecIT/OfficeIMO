namespace OfficeIMO.Drawing;

/// <summary>
/// Shared path command types that OfficeIMO packages can map into their own vector formats.
/// </summary>
public enum OfficePathCommandKind {
    /// <summary>Moves the current point without drawing.</summary>
    MoveTo,

    /// <summary>Draws a straight line to a point.</summary>
    LineTo,

    /// <summary>Draws a cubic Bezier curve using two control points and an end point.</summary>
    CubicBezierTo,

    /// <summary>Closes the current figure.</summary>
    Close
}

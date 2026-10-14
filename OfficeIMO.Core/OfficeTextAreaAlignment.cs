namespace OfficeIMO.Drawing;

/// <summary>Horizontal placement of a paragraph text area inside its padded drawing frame.</summary>
public enum OfficeTextAreaAlignment {
    /// <summary>Use the entire content width for paragraph alignment and wrapping.</summary>
    FullWidth,

    /// <summary>Place the measured paragraph area at the content rectangle's left edge.</summary>
    Left,

    /// <summary>Center the measured paragraph area inside the content rectangle.</summary>
    Center,

    /// <summary>Place the measured paragraph area at the content rectangle's right edge.</summary>
    Right
}

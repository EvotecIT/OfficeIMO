namespace OfficeIMO.Rtf;

/// <summary>
/// Shortcut key metadata stored in an RTF stylesheet <c>{\*\keycode ...}</c> group.
/// </summary>
public sealed class RtfStyleKeyCode {
    /// <summary>Whether the shortcut includes the SHIFT modifier.</summary>
    public bool Shift { get; set; }

    /// <summary>Whether the shortcut includes the CTRL modifier.</summary>
    public bool Control { get; set; }

    /// <summary>Whether the shortcut includes the ALT modifier.</summary>
    public bool Alt { get; set; }

    /// <summary>Optional function-key number from <c>\fn</c>.</summary>
    public int? FunctionKey { get; set; }

    /// <summary>Optional literal key text.</summary>
    public string? Key { get; set; }
}

using System;

namespace OfficeIMO.Drawing;

/// <summary>Semantic palette roles; only roles with native theme slots follow later edits in Office.</summary>
public enum OfficeColorRole {
    /// <summary>A literal RGBA color.</summary>
    None,
    /// <summary>First accent, native accent1.</summary>
    Accent1,
    /// <summary>Second accent, native accent2.</summary>
    Accent2,
    /// <summary>Third accent, native accent3.</summary>
    Accent3,
    /// <summary>Fourth accent, native accent4.</summary>
    Accent4,
    /// <summary>Fifth accent, native accent5.</summary>
    Accent5,
    /// <summary>Sixth accent, native accent6.</summary>
    Accent6,
    /// <summary>Primary text, native dk1.</summary>
    Text,
    /// <summary>Muted text, native dk2.</summary>
    TextMuted,
    /// <summary>Page background, native lt1.</summary>
    Background,
    /// <summary>Positive status, written as a literal.</summary>
    Success,
    /// <summary>Warning status, written as a literal.</summary>
    Warning,
    /// <summary>Danger status, written as a literal.</summary>
    Danger
}

public readonly partial struct OfficeColor {
    private OfficeColor(OfficeColorRole role) {
        _r = _g = _b = 0;
        _a = 255;
        Role = role;
    }
    /// <summary>Gets whether resolution against a theme is required.</summary>
    public bool IsThemeRole => Role != OfficeColorRole.None;
    /// <summary>Gets the native DrawingML slot (accent1-6, dk1, dk2 or lt1), or null for a literal/status role.</summary>
    /// <remarks>Writers map these slot identities to their format's theme representation; dark themes may store light text in dk1.</remarks>
    public string? ThemeSlot => Role switch {
        OfficeColorRole.Accent1 => "accent1", OfficeColorRole.Accent2 => "accent2",
        OfficeColorRole.Accent3 => "accent3", OfficeColorRole.Accent4 => "accent4",
        OfficeColorRole.Accent5 => "accent5", OfficeColorRole.Accent6 => "accent6",
        OfficeColorRole.Text => "dk1", OfficeColorRole.TextMuted => "dk2",
        OfficeColorRole.Background => "lt1", _ => null
    };
    /// <summary>Gets a reference to the first accent.</summary>
    public static OfficeColor Accent1 => new OfficeColor(OfficeColorRole.Accent1);
    /// <summary>Gets a reference to the second accent.</summary>
    public static OfficeColor Accent2 => new OfficeColor(OfficeColorRole.Accent2);
    /// <summary>Gets a reference to the third accent.</summary>
    public static OfficeColor Accent3 => new OfficeColor(OfficeColorRole.Accent3);
    /// <summary>Gets a reference to the fourth accent.</summary>
    public static OfficeColor Accent4 => new OfficeColor(OfficeColorRole.Accent4);
    /// <summary>Gets a reference to the fifth accent.</summary>
    public static OfficeColor Accent5 => new OfficeColor(OfficeColorRole.Accent5);
    /// <summary>Gets a reference to the sixth accent.</summary>
    public static OfficeColor Accent6 => new OfficeColor(OfficeColorRole.Accent6);
    /// <summary>Gets a reference to primary text.</summary>
    public static OfficeColor Text => new OfficeColor(OfficeColorRole.Text);
    /// <summary>Gets a reference to muted text.</summary>
    public static OfficeColor TextMuted => new OfficeColor(OfficeColorRole.TextMuted);
    /// <summary>Gets a reference to the background.</summary>
    public static OfficeColor Background => new OfficeColor(OfficeColorRole.Background);
    /// <summary>Gets a positive-status color resolved at write time.</summary>
    public static OfficeColor Success => new OfficeColor(OfficeColorRole.Success);
    /// <summary>Gets a warning-status color resolved at write time.</summary>
    public static OfficeColor Warning => new OfficeColor(OfficeColorRole.Warning);
    /// <summary>Gets a danger-status color resolved at write time.</summary>
    public static OfficeColor Danger => new OfficeColor(OfficeColorRole.Danger);

    /// <summary>Resolves a role to a literal RGBA color using the supplied theme; literals are returned unchanged.</summary>
    public OfficeColor Resolve(OfficeTheme theme) {
        if (theme == null) throw new ArgumentNullException(nameof(theme));
        OfficeThemePalette palette = theme.Palette;
        return Role switch {
            OfficeColorRole.Accent1 => palette.Accent1, OfficeColorRole.Accent2 => palette.Accent2,
            OfficeColorRole.Accent3 => palette.Accent3, OfficeColorRole.Accent4 => palette.Accent4,
            OfficeColorRole.Accent5 => palette.Accent5, OfficeColorRole.Accent6 => palette.Accent6,
            OfficeColorRole.Text => palette.Text, OfficeColorRole.TextMuted => palette.TextMuted,
            OfficeColorRole.Background => palette.Background, OfficeColorRole.Success => palette.Success,
            OfficeColorRole.Warning => palette.Warning, OfficeColorRole.Danger => palette.Danger, _ => this
        };
    }
    private byte LiteralComponent(byte value) => IsThemeRole
        ? throw new InvalidOperationException("Resolve the color against an OfficeTheme before reading RGBA components or formatting a literal.") : value;
    private static bool TryParseRole(string value, out OfficeColor color) {
        // Enum.TryParse accepts numbers; role parsing deliberately accepts names only.
        foreach (OfficeColorRole role in Enum.GetValues(typeof(OfficeColorRole))) {
            if (role != OfficeColorRole.None && string.Equals(value, role.ToString(), StringComparison.OrdinalIgnoreCase)) {
                color = new OfficeColor(role);
                return true;
            }
        }
        color = default;
        return false;
    }
}

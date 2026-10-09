using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>An explicit native paragraph tab stop with optional text or line paint.</summary>
public sealed class OdfTabStop {
    /// <summary>Creates an absolute tab position with left, center, right or character alignment.</summary>
    /// <remarks>The native type token is left, center, right or char. A char stop requires one Unicode scalar delimiter.</remarks>
    public OdfTabStop(OdfLength position, string type = "left", string? character = null) {
        OfficeTextTabAlignment alignment = type switch {
            "left" => OfficeTextTabAlignment.Left, "center" => OfficeTextTabAlignment.Center,
            "right" => OfficeTextTabAlignment.Right, "char" => OfficeTextTabAlignment.Character,
            _ => throw new ArgumentException("A tab type must be left, center, right or char.", nameof(type))
        };
        if (type == "char" && character == null) throw new ArgumentException("Character tabs require a delimiter.", nameof(character));
        _ = new OfficeTextTabStop(position.ToPoints(), alignment, character ?? ".");
        Position = position; Type = type; Character = character;
    }
    /// <summary>Absolute distance from the native paragraph tab origin.</summary>
    public OdfLength Position { get; }
    /// <summary>Native alignment token.</summary>
    public string Type { get; }
    /// <summary>Character delimiter, or null for non-character stops.</summary>
    public string? Character { get; }
    /// <summary>Optional single Unicode character repeated in the tab gap.</summary>
    public string? LeaderText { get; private set; }
    /// <summary>Optional native line declarations. A declared <see cref="LeaderText"/> takes precedence.</summary>
    public OdfTabLineLeader? LineLeader { get; private set; }
    /// <summary>Referenced text style for a textual leader, or null for the active tab formatting.</summary>
    public OdfStyle? LeaderTextStyle { get; private set; }
    /// <summary>Returns an independent stop with a textual leader; null removes it.</summary>
    /// <remarks>The value must be one non-control Unicode scalar. SPACE retains a blank gap. Projection reports textual glyph layout as approximated.</remarks>
    public OdfTabStop WithLeader(string? text) {
        _ = new OfficeTextTabStop(0).WithLeader(text);
        return new OdfTabStop(Position, Type, Character) { LeaderText = text, LineLeader = LineLeader, LeaderTextStyle = LeaderTextStyle };
    }
    /// <summary>Returns an independent stop with line declarations; null clears them without changing text paint.</summary>
    public OdfTabStop WithLineLeader(OdfTabLineLeader? leader) =>
        new OdfTabStop(Position, Type, Character) { LeaderText = LeaderText, LineLeader = leader, LeaderTextStyle = LeaderTextStyle };
    /// <summary>Returns an independent stop bound to a text style; null clears the binding while retaining the glyph and line declarations.</summary>
    /// <remarks>SetTabStops validates that the style belongs to the paragraph's document and is visible in its package-part scope.</remarks>
    public OdfTabStop WithLeaderTextStyle(OdfStyle? style) {
        if (style != null && style.Family != OdfStyleFamily.Text) throw new ArgumentException("A leader requires a text-family style.", nameof(style));
        return new OdfTabStop(Position, Type, Character) { LeaderText = LeaderText, LineLeader = LineLeader, LeaderTextStyle = style };
    }
    internal XElement ToElement() => new XElement(OdfNamespaces.Style + "tab-stop",
        new XAttribute(OdfNamespaces.Style + "position", Position), new XAttribute(OdfNamespaces.Style + "type", Type),
        Type == "char" ? new XAttribute(OdfNamespaces.Style + "char", Character!) : null,
        LeaderText != null ? new XAttribute(OdfNamespaces.Style + "leader-text", LeaderText) : null,
        LeaderTextStyle != null ? new XAttribute(OdfNamespaces.Style + "leader-text-style", LeaderTextStyle.Name) : null,
        // Native Draw drops a text-only leader when the line style defaults to none.
        // Text takes precedence over this explicit activating line declaration.
        LineLeader != null ? new XAttribute(OdfNamespaces.Style + "leader-style", LineLeader.Style) :
            LeaderText != null ? new XAttribute(OdfNamespaces.Style + "leader-style", string.IsNullOrWhiteSpace(LeaderText) ? "none" : "solid") : null,
        LineLeader != null ? new XAttribute(OdfNamespaces.Style + "leader-type", LineLeader.Type) : null,
        LineLeader != null ? new XAttribute(OdfNamespaces.Style + "leader-width", LineLeader.Width) : null,
        LineLeader != null ? new XAttribute(OdfNamespaces.Style + "leader-color", LineLeader.Color?.ToString() ?? "font-color") : null);
}

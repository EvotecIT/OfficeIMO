using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Literal colors defining a document theme palette.</summary>
public sealed class OfficeThemePalette {
    /// <summary>Gets the literal Background color.</summary>
    public OfficeColor Background { get; }
    /// <summary>Gets the literal Text color.</summary>
    public OfficeColor Text { get; }
    /// <summary>Gets the literal TextMuted color.</summary>
    public OfficeColor TextMuted { get; }
    /// <summary>Gets the literal Accent1 color.</summary>
    public OfficeColor Accent1 { get; }
    /// <summary>Gets the literal Accent2 color.</summary>
    public OfficeColor Accent2 { get; }
    /// <summary>Gets the literal Accent3 color.</summary>
    public OfficeColor Accent3 { get; }
    /// <summary>Gets the literal Accent4 color.</summary>
    public OfficeColor Accent4 { get; }
    /// <summary>Gets the literal Accent5 color.</summary>
    public OfficeColor Accent5 { get; }
    /// <summary>Gets the literal Accent6 color.</summary>
    public OfficeColor Accent6 { get; }
    /// <summary>Gets the literal Success color.</summary>
    public OfficeColor Success { get; }
    /// <summary>Gets the literal Warning color.</summary>
    public OfficeColor Warning { get; }
    /// <summary>Gets the literal Danger color.</summary>
    public OfficeColor Danger { get; }

    /// <summary>Creates immutable literal colors defining a document theme palette.</summary>
    public OfficeThemePalette(OfficeColor background, OfficeColor text, OfficeColor textMuted, OfficeColor accent1, OfficeColor accent2, OfficeColor accent3, OfficeColor accent4, OfficeColor accent5, OfficeColor accent6, OfficeColor success, OfficeColor warning, OfficeColor danger) {
        ThemeValues.Literal(background, nameof(background));
        ThemeValues.Literal(text, nameof(text));
        ThemeValues.Literal(textMuted, nameof(textMuted));
        ThemeValues.Literal(accent1, nameof(accent1));
        ThemeValues.Literal(accent2, nameof(accent2));
        ThemeValues.Literal(accent3, nameof(accent3));
        ThemeValues.Literal(accent4, nameof(accent4));
        ThemeValues.Literal(accent5, nameof(accent5));
        ThemeValues.Literal(accent6, nameof(accent6));
        ThemeValues.Literal(success, nameof(success));
        ThemeValues.Literal(warning, nameof(warning));
        ThemeValues.Literal(danger, nameof(danger));
        Background = background;
        Text = text;
        TextMuted = textMuted;
        Accent1 = accent1;
        Accent2 = accent2;
        Accent3 = accent3;
        Accent4 = accent4;
        Accent5 = accent5;
        Accent6 = accent6;
        Success = success;
        Warning = warning;
        Danger = danger;
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemePalette With(OfficeColor? background = null, OfficeColor? text = null, OfficeColor? textMuted = null, OfficeColor? accent1 = null, OfficeColor? accent2 = null, OfficeColor? accent3 = null, OfficeColor? accent4 = null, OfficeColor? accent5 = null, OfficeColor? accent6 = null, OfficeColor? success = null, OfficeColor? warning = null, OfficeColor? danger = null) =>
        new OfficeThemePalette(background ?? Background, text ?? Text, textMuted ?? TextMuted, accent1 ?? Accent1, accent2 ?? Accent2, accent3 ?? Accent3, accent4 ?? Accent4, accent5 ?? Accent5, accent6 ?? Accent6, success ?? Success, warning ?? Warning, danger ?? Danger);
}

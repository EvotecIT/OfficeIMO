namespace OfficeIMO.MarkdownRenderer.Wpf;

/// <summary>
/// Built-in renderer presets exposed by the WPF markdown host control.
/// </summary>
public enum MarkdownViewPreset {
    /// <summary>
    /// Uses the strict generic renderer preset.
    /// </summary>
    Strict = 0,

    /// <summary>
    /// Uses the strict minimal preset with optional client-side helpers disabled.
    /// </summary>
    StrictMinimal = 1,

    /// <summary>
    /// Uses the relaxed preset intended for trusted or controlled markdown content.
    /// </summary>
    Relaxed = 2
}

namespace OfficeIMO.Markdown;

/// <summary>
/// Built-in accent color schemes that can be applied to shared Markdown visual themes.
/// </summary>
public enum MarkdownColorSchemeKind {
    /// <summary>Use the theme's default blue/slate document palette.</summary>
    Default,

    /// <summary>Blue accent palette for general documents.</summary>
    Blue,

    /// <summary>Emerald accent palette for calm operational reports.</summary>
    Emerald,

    /// <summary>Indigo accent palette for technical documents.</summary>
    Indigo,

    /// <summary>Rose accent palette for review or exception reports.</summary>
    Rose,

    /// <summary>Amber accent palette for warnings and operational notes.</summary>
    Amber,

    /// <summary>Slate accent palette for understated internal documents.</summary>
    Slate
}

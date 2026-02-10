namespace OfficeIMO.Markdown;

/// <summary>
/// Built-in style presets for rendered HTML.
/// </summary>
public enum HtmlStyle {
    /// <summary>No styling (only browser defaults).</summary>
    Plain,
    /// <summary>Lightweight readable defaults.</summary>
    Clean,
    /// <summary>GitHub-like light theme.</summary>
    GithubLight,
    /// <summary>GitHub-like dark theme.</summary>
    GithubDark,
    /// <summary>Auto light/dark using prefers-color-scheme.</summary>
    GithubAuto,
    /// <summary>Compact chat-friendly light theme (intended for embedding inside chat UIs).</summary>
    ChatLight,
    /// <summary>Compact chat-friendly dark theme (intended for embedding inside chat UIs).</summary>
    ChatDark,
    /// <summary>Compact chat-friendly auto theme (prefers-color-scheme + data-theme overrides).</summary>
    ChatAuto,
    /// <summary>Word-like styling: Calibri/Cambria headings, comfortable spacing, and Word-ish tables/lists.</summary>
    Word
}

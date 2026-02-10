namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// How the renderer should behave when guardrail limits are exceeded.
/// </summary>
public enum OverflowHandling {
    /// <summary>Truncate the input when it is safe to do so (Markdown). For HTML overflow, falls back to RenderError.</summary>
    Truncate,
    /// <summary>Render a small in-band warning message instead of the full content.</summary>
    RenderError,
    /// <summary>Throw an exception.</summary>
    Throw
}


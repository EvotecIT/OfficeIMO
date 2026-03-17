namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Supported fenced metadata value kinds for plugin option schemas.
/// </summary>
public enum MarkdownFenceOptionValueKind {
    /// <summary>Plain string metadata.</summary>
    String,

    /// <summary>Boolean metadata such as <c>pinned</c> or <c>compact=false</c>.</summary>
    Boolean,

    /// <summary>32-bit integer metadata such as <c>maxItems=12</c>.</summary>
    Int32
}

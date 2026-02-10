namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Optional hint for chat-style HTML wrappers (bubble alignment/colors).
/// </summary>
public enum ChatMessageRole {
    /// <summary>Assistant/bot output.</summary>
    Assistant,
    /// <summary>User input.</summary>
    User,
    /// <summary>System/status messages (e.g. connection state).</summary>
    System
}

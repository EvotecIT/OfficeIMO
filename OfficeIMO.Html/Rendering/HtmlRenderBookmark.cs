namespace OfficeIMO.Html;

/// <summary>Expansion preference retained for a rendered document outline entry.</summary>
public enum HtmlRenderBookmarkState {
    /// <summary>Uses the PDF document's default outline expansion policy.</summary>
    Default,
    /// <summary>Requests an expanded outline entry.</summary>
    Open,
    /// <summary>Requests a collapsed outline entry.</summary>
    Closed
}

internal sealed class HtmlRenderBookmarkDefinition {
    internal HtmlRenderBookmarkDefinition(int level, string? label, HtmlRenderBookmarkState state, bool suppressed, int sourceOrder) {
        Level = level;
        Label = label;
        State = state;
        Suppressed = suppressed;
        SourceOrder = sourceOrder;
    }

    internal int Level { get; }
    internal string? Label { get; }
    internal HtmlRenderBookmarkState State { get; }
    internal bool Suppressed { get; }
    internal int SourceOrder { get; }
}

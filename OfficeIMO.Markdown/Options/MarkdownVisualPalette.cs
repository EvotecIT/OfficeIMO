namespace OfficeIMO.Markdown;

/// <summary>
/// Shared semantic color palette used by Markdown exporters.
/// </summary>
public sealed class MarkdownVisualPalette {
    /// <summary>Primary accent color for headings, links, rules, and active UI affordances.</summary>
    public MarkdownColor Accent { get; set; } = MarkdownColor.Parse("#2563EB");

    /// <summary>Document heading color.</summary>
    public MarkdownColor Heading { get; set; } = MarkdownColor.Parse("#111827");

    /// <summary>Main body text color.</summary>
    public MarkdownColor Text { get; set; } = MarkdownColor.Parse("#1F2937");

    /// <summary>Secondary text color for metadata, captions, and muted content.</summary>
    public MarkdownColor MutedText { get; set; } = MarkdownColor.Parse("#64748B");

    /// <summary>Page or body background color.</summary>
    public MarkdownColor Background { get; set; } = MarkdownColor.Parse("#FFFFFF");

    /// <summary>Panel, code, and callout surface color.</summary>
    public MarkdownColor Surface { get; set; } = MarkdownColor.Parse("#F8FAFC");

    /// <summary>Default border and rule color.</summary>
    public MarkdownColor Border { get; set; } = MarkdownColor.Parse("#CBD5E1");

    /// <summary>Code block background color.</summary>
    public MarkdownColor CodeBackground { get; set; } = MarkdownColor.Parse("#F6F8FA");

    /// <summary>Table header background color.</summary>
    public MarkdownColor TableHeaderBackground { get; set; } = MarkdownColor.Parse("#EFF6FF");

    /// <summary>Table header text color.</summary>
    public MarkdownColor TableHeaderText { get; set; } = MarkdownColor.Parse("#0F172A");

    /// <summary>Alternating table row background color.</summary>
    public MarkdownColor TableStripeBackground { get; set; } = MarkdownColor.Parse("#F8FAFC");

    /// <summary>Creates a copy of this palette.</summary>
    public MarkdownVisualPalette Clone() => new MarkdownVisualPalette {
        Accent = Accent,
        Heading = Heading,
        Text = Text,
        MutedText = MutedText,
        Background = Background,
        Surface = Surface,
        Border = Border,
        CodeBackground = CodeBackground,
        TableHeaderBackground = TableHeaderBackground,
        TableHeaderText = TableHeaderText,
        TableStripeBackground = TableStripeBackground
    };
}

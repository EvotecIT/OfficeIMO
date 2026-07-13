namespace OfficeIMO.Markdown;

/// <summary>
/// Shared semantic color palette used by Markdown exporters.
/// </summary>
public sealed class MarkdownVisualPalette {
    /// <summary>Primary accent color for headings, links, rules, and active UI affordances.</summary>
    public OfficeColor Accent { get; set; } = OfficeColor.Parse("#2563EB");

    /// <summary>Document heading color.</summary>
    public OfficeColor Heading { get; set; } = OfficeColor.Parse("#111827");

    /// <summary>Main body text color.</summary>
    public OfficeColor Text { get; set; } = OfficeColor.Parse("#1F2937");

    /// <summary>Secondary text color for metadata, captions, and muted content.</summary>
    public OfficeColor MutedText { get; set; } = OfficeColor.Parse("#64748B");

    /// <summary>Page or body background color.</summary>
    public OfficeColor Background { get; set; } = OfficeColor.Parse("#FFFFFF");

    /// <summary>Panel, code, and callout surface color.</summary>
    public OfficeColor Surface { get; set; } = OfficeColor.Parse("#F8FAFC");

    /// <summary>Default border and rule color.</summary>
    public OfficeColor Border { get; set; } = OfficeColor.Parse("#CBD5E1");

    /// <summary>Code block background color.</summary>
    public OfficeColor CodeBackground { get; set; } = OfficeColor.Parse("#F6F8FA");

    /// <summary>Table header background color.</summary>
    public OfficeColor TableHeaderBackground { get; set; } = OfficeColor.Parse("#EFF6FF");

    /// <summary>Table header text color.</summary>
    public OfficeColor TableHeaderText { get; set; } = OfficeColor.Parse("#0F172A");

    /// <summary>Alternating table row background color.</summary>
    public OfficeColor TableStripeBackground { get; set; } = OfficeColor.Parse("#F8FAFC");

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

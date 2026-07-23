namespace OfficeIMO.Markdown;

/// <summary>
/// Marker family used by an ordered list.
/// </summary>
public enum MarkdownOrderedListMarkerStyle {
    /// <summary>Decimal markers such as <c>1.</c>.</summary>
    Decimal,
    /// <summary>Lowercase alphabetic markers such as <c>a.</c>.</summary>
    LowerAlpha,
    /// <summary>Uppercase alphabetic markers such as <c>A.</c>.</summary>
    UpperAlpha,
    /// <summary>Lowercase roman markers such as <c>iv.</c>.</summary>
    LowerRoman,
    /// <summary>Uppercase roman markers such as <c>IV.</c>.</summary>
    UpperRoman
}

/// <summary>
/// Ordered (numbered) list.
/// </summary>
public sealed class OrderedListBlock : MarkdownBlock, IMarkdownListBlock, ISyntaxMarkdownBlock {
    /// <summary>Items within the ordered list.</summary>
    public List<ListItem> Items { get; } = new List<ListItem>();
    /// <summary>Starting number (default 1).</summary>
    public int Start { get; set; } = 1;
    /// <summary>Whether top-level item numbering descends from <see cref="Start"/>.</summary>
    public bool Reversed { get; set; }
    /// <summary>Ordered list marker family.</summary>
    public MarkdownOrderedListMarkerStyle MarkerStyle { get; set; } = MarkdownOrderedListMarkerStyle.Decimal;
    /// <summary>Marker delimiter used when writing generated list markers.</summary>
    public char MarkerDelimiter { get; set; } = '.';

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() =>
        MarkdownListRendering.RenderMarkdown(
            Attributes,
            Items,
            (item, topLevelIndex) => {
                string markerText = item.MarkerText ?? (item.Level == 0
                    ? FormatMarker(Reversed ? Start - topLevelIndex : Start + topLevelIndex, MarkerStyle, MarkerDelimiter)
                    : FormatMarker(1, MarkerStyle, MarkerDelimiter));
                string baseMarker = markerText + " ";
                return item.IsTask
                    ? baseMarker + "[" + (item.Checked ? "x" : " ") + "] "
                    : baseMarker;
            });

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() =>
        RenderHtml(renderItemAttributes: false);

    internal string RenderHtml(bool renderItemAttributes) =>
        MarkdownListRendering.RenderHtml(
            "ol",
            Items,
            Attributes,
            _ => RenderOrderedListAttributes(),
            renderItemAttributes);

    IReadOnlyList<ListItem> IMarkdownListBlock.ListItems => Items;
    MarkdownSyntaxKind IMarkdownListBlock.ListSyntaxKind => MarkdownSyntaxKind.OrderedList;
    string? IMarkdownListBlock.ListLiteral => Start.ToString(System.Globalization.CultureInfo.InvariantCulture);
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownListSyntax.BuildListBlockNode(this, span);

    private string RenderOrderedListAttributes() {
        var attributes = new System.Text.StringBuilder();
        var type = MarkerStyle switch {
            MarkdownOrderedListMarkerStyle.LowerAlpha => "a",
            MarkdownOrderedListMarkerStyle.UpperAlpha => "A",
            MarkdownOrderedListMarkerStyle.LowerRoman => "i",
            MarkdownOrderedListMarkerStyle.UpperRoman => "I",
            _ => null
        };
        if (type != null) {
            attributes.Append(" type=\"").Append(type).Append('"');
        }

        if (Start != 1 || Reversed) {
            attributes.Append(" start=\"").Append(Start.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append('"');
        }

        if (Reversed) {
            attributes.Append(" reversed");
        }

        return attributes.ToString();
    }

    internal static string FormatMarker(int value, MarkdownOrderedListMarkerStyle style, char delimiter) {
        var text = style switch {
            MarkdownOrderedListMarkerStyle.LowerAlpha => FormatAlpha(value, uppercase: false),
            MarkdownOrderedListMarkerStyle.UpperAlpha => FormatAlpha(value, uppercase: true),
            MarkdownOrderedListMarkerStyle.LowerRoman => FormatRoman(value, uppercase: false),
            MarkdownOrderedListMarkerStyle.UpperRoman => FormatRoman(value, uppercase: true),
            _ => value.ToString(System.Globalization.CultureInfo.InvariantCulture)
        };

        return text + delimiter;
    }

    private static string FormatAlpha(int value, bool uppercase) {
        if (value < 1 || value > 26) {
            return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        var ch = (char)((uppercase ? 'A' : 'a') + value - 1);
        return ch.ToString();
    }

    private static string FormatRoman(int value, bool uppercase) {
        if (value < 1 || value > 39) {
            return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        var tens = value / 10;
        var ones = value % 10;
        var text = new string('x', tens) + (ones switch {
            0 => string.Empty,
            1 => "i",
            2 => "ii",
            3 => "iii",
            4 => "iv",
            5 => "v",
            6 => "vi",
            7 => "vii",
            8 => "viii",
            9 => "ix",
            _ => string.Empty
        });

        return uppercase ? text.ToUpperInvariant() : text;
    }
}

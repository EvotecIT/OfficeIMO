using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Renders <see cref="TextProperty"/> into a TextBlock with the query terms emphasized.</summary>
public static class StudioTextHighlight {
    public static readonly AttachedProperty<string?> TextProperty =
        AvaloniaProperty.RegisterAttached<TextBlock, string?>("Text", typeof(StudioTextHighlight));

    public static readonly AttachedProperty<string?> QueryProperty =
        AvaloniaProperty.RegisterAttached<TextBlock, string?>("Query", typeof(StudioTextHighlight));

    public static readonly AttachedProperty<IBrush?> HighlightBrushProperty =
        AvaloniaProperty.RegisterAttached<TextBlock, IBrush?>("HighlightBrush", typeof(StudioTextHighlight));

    static StudioTextHighlight() {
        TextProperty.Changed.AddClassHandler<TextBlock>((block, _) => Update(block));
        QueryProperty.Changed.AddClassHandler<TextBlock>((block, _) => Update(block));
        HighlightBrushProperty.Changed.AddClassHandler<TextBlock>((block, _) => Update(block));
    }

    public static string? GetText(TextBlock element) => element.GetValue(TextProperty);
    public static void SetText(TextBlock element, string? value) => element.SetValue(TextProperty, value);
    public static string? GetQuery(TextBlock element) => element.GetValue(QueryProperty);
    public static void SetQuery(TextBlock element, string? value) => element.SetValue(QueryProperty, value);
    public static IBrush? GetHighlightBrush(TextBlock element) => element.GetValue(HighlightBrushProperty);
    public static void SetHighlightBrush(TextBlock element, IBrush? value) => element.SetValue(HighlightBrushProperty, value);

    /// <summary>Returns non-overlapping (start, length) ranges of <paramref name="text"/> that match any query term.</summary>
    internal static IReadOnlyList<(int Start, int Length)> FindMatches(string text, string? query) {
        var ranges = new List<(int Start, int Length)>();
        if (string.IsNullOrEmpty(text) || string.IsNullOrWhiteSpace(query)) return ranges;
        var covered = new bool[text.Length];
        foreach (string term in query.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) {
            int index = text.IndexOf(term, StringComparison.CurrentCultureIgnoreCase);
            if (index < 0) continue;
            for (int offset = index; offset < index + term.Length && offset < covered.Length; offset++) covered[offset] = true;
        }
        int start = -1;
        for (int index = 0; index <= covered.Length; index++) {
            bool on = index < covered.Length && covered[index];
            if (on && start < 0) start = index;
            if (!on && start >= 0) { ranges.Add((start, index - start)); start = -1; }
        }
        return ranges;
    }

    private static void Update(TextBlock block) {
        string text = GetText(block) ?? string.Empty;
        IReadOnlyList<(int Start, int Length)> matches = FindMatches(text, GetQuery(block));
        if (matches.Count == 0) {
            block.Inlines = null;
            block.Text = text;
            return;
        }
        var inlines = new InlineCollection();
        int position = 0;
        foreach ((int start, int length) in matches) {
            if (start > position) inlines.Add(new Run(text[position..start]));
            var run = new Run(text.Substring(start, length)) { FontWeight = FontWeight.Bold };
            if (GetHighlightBrush(block) is { } brush) run.Foreground = brush;
            inlines.Add(run);
            position = start + length;
        }
        if (position < text.Length) inlines.Add(new Run(text[position..]));
        block.Inlines = inlines;
    }
}

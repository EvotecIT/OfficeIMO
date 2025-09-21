namespace OfficeIMO.Markdown;

/// <summary>
/// Builder for block quotes supporting raw lines or nested blocks.
/// </summary>
public sealed class QuoteBuilder {
    private readonly System.Collections.Generic.List<Entry> _entries = new();

    /// <summary>Adds a raw text line to the quote.</summary>
    public QuoteBuilder Line(string text) {
        _entries.Add(new Entry(text ?? string.Empty));
        return this;
    }

    /// <summary>Adds multiple raw lines to the quote.</summary>
    public QuoteBuilder Lines(System.Collections.Generic.IEnumerable<string> lines) {
        if (lines == null) return this;
        foreach (var line in lines) Line(line ?? string.Empty);
        return this;
    }

    /// <summary>Adds a paragraph with the specified text.</summary>
    public QuoteBuilder P(string text) {
        return Block(new ParagraphBlock(new InlineSequence().Text(text ?? string.Empty)));
    }

    /// <summary>Adds a paragraph built via <see cref="ParagraphBuilder"/>.</summary>
    public QuoteBuilder P(System.Action<ParagraphBuilder> build) {
        if (build == null) return this;
        ParagraphBuilder builder = new ParagraphBuilder();
        build(builder);
        return Block(new ParagraphBlock(builder.Inlines));
    }

    /// <summary>Adds a nested block quote.</summary>
    public QuoteBuilder Quote(System.Action<QuoteBuilder> build) {
        if (build == null) return this;
        QuoteBuilder nested = new QuoteBuilder();
        build(nested);
        return Block(nested.Build());
    }

    /// <summary>Adds an arbitrary markdown block inside the quote.</summary>
    public QuoteBuilder Block(IMarkdownBlock block) {
        if (block == null) throw new System.ArgumentNullException(nameof(block));
        _entries.Add(new Entry(block));
        return this;
    }

    internal QuoteBlock Build() {
        bool hasBlocks = false;
        foreach (var entry in _entries) {
            if (entry.Block != null) { hasBlocks = true; break; }
        }

        QuoteBlock quote = new QuoteBlock();
        if (!hasBlocks) {
            foreach (var entry in _entries) quote.Lines.AddRange(SplitLines(entry.Line));
            return quote;
        }

        System.Collections.Generic.List<string> buffer = new();
        foreach (var entry in _entries) {
            if (entry.Block != null) {
                FlushBuffer(buffer, quote);
                quote.Children.Add(entry.Block);
            } else {
                foreach (var part in SplitLines(entry.Line)) buffer.Add(part);
            }
        }
        FlushBuffer(buffer, quote);
        return quote;
    }

    private static System.Collections.Generic.IEnumerable<string> SplitLines(string? value) {
        if (value == null) yield break;
        string normalized = value.Replace("\r\n", "\n").Replace('\r', '\n');
        int start = 0;
        for (int i = 0; i < normalized.Length; i++) {
            if (normalized[i] == '\n') {
                yield return normalized.Substring(start, i - start);
                start = i + 1;
            }
        }
        if (start <= normalized.Length) yield return normalized.Substring(start);
    }

    private static void FlushBuffer(System.Collections.Generic.List<string> buffer, QuoteBlock quote) {
        if (buffer.Count == 0) return;
        foreach (var paragraph in BuildParagraphs(buffer)) quote.Children.Add(paragraph);
        buffer.Clear();
    }

    private static System.Collections.Generic.IEnumerable<ParagraphBlock> BuildParagraphs(System.Collections.Generic.List<string> lines) {
        System.Collections.Generic.List<string> current = new();
        foreach (var line in lines) {
            if (line.Length == 0) {
                if (current.Count > 0) {
                    yield return CreateParagraph(current);
                    current.Clear();
                }
                continue;
            }
            current.Add(line);
        }
        if (current.Count > 0) {
            yield return CreateParagraph(current);
            current.Clear();
        }
    }

    private static ParagraphBlock CreateParagraph(System.Collections.Generic.List<string> lines) {
        InlineSequence sequence = new InlineSequence();
        for (int i = 0; i < lines.Count; i++) {
            if (i > 0) sequence.HardBreak();
            sequence.Text(lines[i]);
        }
        return new ParagraphBlock(sequence);
    }

    private readonly struct Entry {
        public string? Line { get; }
        public IMarkdownBlock? Block { get; }
        public Entry(string line) { Line = line; Block = null; }
        public Entry(IMarkdownBlock block) { Line = null; Block = block; }
    }
}

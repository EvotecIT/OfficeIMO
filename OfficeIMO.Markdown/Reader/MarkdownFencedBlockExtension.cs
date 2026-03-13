namespace OfficeIMO.Markdown;

/// <summary>
/// Context passed to fenced block extension factories during parsing.
/// </summary>
public sealed class MarkdownFencedBlockFactoryContext {
    internal MarkdownFencedBlockFactoryContext(string language, string content, bool isFenced, string? caption) {
        Language = language ?? string.Empty;
        Content = content ?? string.Empty;
        IsFenced = isFenced;
        Caption = caption;
    }

    /// <summary>Fence language / info string that opened the block.</summary>
    public string Language { get; }

    /// <summary>Raw fenced block payload.</summary>
    public string Content { get; }

    /// <summary>Whether the source block was fenced rather than indented.</summary>
    public bool IsFenced { get; }

    /// <summary>Optional caption parsed immediately after the block.</summary>
    public string? Caption { get; }
}

/// <summary>
/// Creates a specialized markdown block for a matching fenced code block during parsing.
/// Returning <see langword="null"/> falls back to the standard <see cref="CodeBlock"/>.
/// </summary>
public delegate IMarkdownBlock? MarkdownFencedBlockFactory(MarkdownFencedBlockFactoryContext context);

/// <summary>
/// Defines a language-based fenced block parser extension.
/// </summary>
public sealed class MarkdownFencedBlockExtension {
    /// <summary>
    /// Creates a new fenced block extension.
    /// </summary>
    public MarkdownFencedBlockExtension(string name, IEnumerable<string> languages, MarkdownFencedBlockFactory createBlock) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name is required.", nameof(name));
        }

        if (languages == null) {
            throw new ArgumentNullException(nameof(languages));
        }

        CreateBlock = createBlock ?? throw new ArgumentNullException(nameof(createBlock));

        var normalized = new List<string>();
        foreach (var language in languages) {
            var value = (language ?? string.Empty).Trim();
            if (value.Length == 0) {
                continue;
            }

            var exists = false;
            for (int i = 0; i < normalized.Count; i++) {
                if (string.Equals(normalized[i], value, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                normalized.Add(value);
            }
        }

        if (normalized.Count == 0) {
            throw new ArgumentException("At least one fenced block language is required.", nameof(languages));
        }

        Name = name.Trim();
        Languages = normalized;
    }

    /// <summary>Friendly extension name used for diagnostics and documentation.</summary>
    public string Name { get; }

    /// <summary>Fence languages handled by this extension.</summary>
    public IReadOnlyList<string> Languages { get; }

    /// <summary>Factory invoked for matching fenced blocks during parsing.</summary>
    public MarkdownFencedBlockFactory CreateBlock { get; }
}

namespace OfficeIMO.Yaml;

/// <summary>
/// Rewrites YAML in one consistent layout. It writes the document through the same emitter as
/// <see cref="YamlSaveOptions.Reformat"/>, so values, scalar spellings, anchors, aliases, tags, comments and the blank
/// lines above entries are kept, and formatting formatted text returns it unchanged.
/// </summary>
/// <remarks>
/// The layout is the emitter's: block collections indented by the configured widths, flow collections on one line,
/// multi-line plain and folded scalars laid out again. Comments inside flow collections are dropped, as with
/// <see cref="YamlSaveOptions.Reformat"/>. To repair individual problems while keeping every other line as it is,
/// use <see cref="YamlLinter.Fix(string, YamlLintOptions?)"/> instead.
/// </remarks>
public static class YamlFormatter {
    /// <summary>Formats YAML text.</summary>
    /// <param name="yaml">The text to format.</param>
    /// <param name="options">The layout, or <see langword="null"/> to keep the document's own conventions.</param>
    /// <param name="loadOptions">How the text is loaded, or <see langword="null"/> for the defaults.</param>
    /// <returns>The formatted text.</returns>
    /// <exception cref="YamlException">The text is not valid YAML.</exception>
    /// <exception cref="NotSupportedException">The document was loaded with <see cref="YamlDialect.Unity"/>. Unity reads
    /// only the layout it writes, so its files are repaired with <see cref="YamlLinter.Fix(string, YamlLintOptions?)"/>,
    /// which changes only the affected lines.</exception>
    public static string Format(string yaml, YamlFormatOptions? options = null, YamlLoadOptions? loadOptions = null) {
        if (yaml is null) throw new ArgumentNullException(nameof(yaml));
        using YamlDocument document = YamlDocument.Parse(yaml, loadOptions);
        return Format(document, options, loadOptions);
    }

    /// <summary>Formats a document, including any edits made to it, without changing the document.</summary>
    /// <param name="document">The document to format.</param>
    /// <param name="options">The layout, or <see langword="null"/> to keep the document's own conventions.</param>
    /// <returns>The formatted text.</returns>
    /// <exception cref="NotSupportedException">The document was loaded with <see cref="YamlDialect.Unity"/>. Unity reads
    /// only the layout it writes, so its files are repaired with <see cref="YamlLinter.Fix(string, YamlLintOptions?)"/>,
    /// which changes only the affected lines.</exception>
    public static string Format(YamlDocument document, YamlFormatOptions? options) => Format(document, options, null);

    private static string Format(YamlDocument document, YamlFormatOptions? options, YamlLoadOptions? loadOptions) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (document.Dialect == YamlDialect.Unity) {
            throw new NotSupportedException("A Unity file cannot be formatted: Unity reads only the layout it writes. Use YamlLinter.Fix to repair it, which changes only the affected lines.");
        }
        options ??= new YamlFormatOptions();
        var save = new YamlSaveOptions {
            Reformat = true,
            WriterOptions = options.ToWriterOptions(document.WriterOptions),
            QuoteStyle = options.QuoteStyle
        };
        string text = document.ToString(save);
        if (options.FinalNewline) return text;
        string trimmed = YamlModelWriter.TrimFinalBreak(text);
        if (trimmed.Length == text.Length) return text;
        // A block scalar at the end keeps its final line break in its value; leaving the break out would change it.
        return YamlValueEquivalence.SameValues(text, trimmed, loadOptions) ? trimmed : text;
    }
}

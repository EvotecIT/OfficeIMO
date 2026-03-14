namespace OfficeIMO.Markdown;

/// <summary>
/// Helper methods for registering OfficeIMO.Markdown's built-in block syntax extensions onto a reader profile.
/// </summary>
public static class MarkdownReaderBuiltInExtensions {
    /// <summary>Stable registration name for Docs-style callouts.</summary>
    public const string CalloutsExtensionName = "OfficeIMO.Callouts";

    /// <summary>Stable registration name for TOC placeholder blocks.</summary>
    public const string TocPlaceholdersExtensionName = "OfficeIMO.TocPlaceholders";

    /// <summary>Stable registration name for footnote definition blocks.</summary>
    public const string FootnotesExtensionName = "OfficeIMO.Footnotes";

    /// <summary>
    /// Registers the OfficeIMO default built-in block extensions: callouts, TOC placeholders, and footnotes.
    /// </summary>
    public static void RegisterOfficeIMODefaults(MarkdownReaderOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddCallouts(options);
        AddTocPlaceholders(options);
        AddFootnotes(options);
    }

    /// <summary>Adds Docs-style callout parsing to the supplied reader options.</summary>
    public static void AddCallouts(MarkdownReaderOptions options) {
        AddIfMissing(
            options,
            CalloutsExtensionName,
            MarkdownBlockParserPlacement.AfterFrontMatter,
            new MarkdownReader.CalloutParser(),
            static readerOptions => readerOptions.Callouts);
    }

    /// <summary>Adds TOC placeholder parsing to the supplied reader options.</summary>
    public static void AddTocPlaceholders(MarkdownReaderOptions options) {
        AddIfMissing(
            options,
            TocPlaceholdersExtensionName,
            MarkdownBlockParserPlacement.AfterHtmlBlocks,
            new MarkdownReader.TocParser(),
            static readerOptions => readerOptions.TocPlaceholders);
    }

    /// <summary>Adds footnote definition parsing to the supplied reader options.</summary>
    public static void AddFootnotes(MarkdownReaderOptions options) {
        AddIfMissing(
            options,
            FootnotesExtensionName,
            MarkdownBlockParserPlacement.AfterReferenceLinkDefinitions,
            new MarkdownReader.FootnoteParser(),
            static readerOptions => readerOptions.Footnotes);
    }

    /// <summary>Returns <see langword="true"/> when a named built-in extension is already registered.</summary>
    public static bool HasRegistration(MarkdownReaderOptions options, string extensionName) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (string.IsNullOrWhiteSpace(extensionName)) {
            return false;
        }

        return options.BlockParserExtensions.Any(extension =>
            extension != null
            && string.Equals(extension.Name, extensionName, StringComparison.OrdinalIgnoreCase));
    }

    private static void AddIfMissing(
        MarkdownReaderOptions options,
        string name,
        MarkdownBlockParserPlacement placement,
        IMarkdownBlockParser parser,
        Func<MarkdownReaderOptions, bool> isEnabled) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (HasRegistration(options, name)) {
            return;
        }

        options.BlockParserExtensions.Add(new MarkdownBlockParserExtension(name, placement, parser, isEnabled));
    }
}

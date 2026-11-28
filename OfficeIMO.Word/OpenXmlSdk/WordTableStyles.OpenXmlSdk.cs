using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

// The SDK side of the built-in table styles: the definitions the SDK path adds to styles parts, until styles move to
// the owned model.
internal static partial class WordTableStyles {
    /// <summary>The <see cref="Style"/> definition of a built-in table style, ready to add to a styles part.</summary>
    internal static Style GetStyleDefinition(WordTableStyle style) {
        if (style == WordTableStyle.OfficeIMOModern) return WordThemeDefaults.CreateTableStyle(OfficeTheme.Default);
        return (Style)ParsedDefinitions.GetOrAdd(style, key => new Lazy<Style>(() => ParseDefinition(key))).Value.CloneNode(true);
    }

    // Each definition is parsed once, even when threads ask for it together; callers get copies they can add to a document.
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<WordTableStyle, Lazy<Style>> ParsedDefinitions = new();

    private static Style ParseDefinition(WordTableStyle style) {
        var definition = new Style(GetStyleMarkup(style)!);
        // The styles part declares the namespace; the definition keeps none of its own.
        definition.RemoveNamespaceDeclaration("w");
        _ = definition.OuterXml;
        return definition;
    }
}

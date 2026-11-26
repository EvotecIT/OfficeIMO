using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

// The SDK side of the built-in table styles: the elements the SDK path adds to documents.
internal static partial class WordTableStyles {
    /// <summary>Creates a <see cref="TableStyle"/> element that applies the specified table style.</summary>
    internal static TableStyle GetStyle(WordTableStyle style) => new TableStyle { Val = GetStyleId(style) };

    /// <summary>Whether the styles part already defines the table style.</summary>
    internal static bool IsAvailableStyle(Styles styles, WordTableStyle style) {
        string styleId = GetStyleId(style);
        foreach (Style current in styles.OfType<Style>()) {
            if (current.StyleId == styleId) return true;
        }
        return false;
    }

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

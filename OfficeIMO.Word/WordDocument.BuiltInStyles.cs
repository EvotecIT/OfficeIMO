using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private BuiltInStyleIndex? _builtInStyleIndex;

    /// <summary>
    /// Finds this document's identifier for a built-in style: the style whose <c>w:name</c> is the built-in name,
    /// whatever its localized identifier, otherwise a style with the English identifier.
    /// </summary>
    /// <param name="builtInStyleId">The English built-in identifier, such as <c>Heading1</c> or <c>Quote</c>.</param>
    /// <returns>The document's identifier, or <see langword="null"/> when the document defines neither.</returns>
    internal string? FindBuiltInStyleId(string builtInStyleId) {
        if (GetUnparsedDefaultPackageStyles() is { } defaults) {
            return defaults.Dependencies.ContainsKey(builtInStyleId) ? builtInStyleId : null;
        }
        BuiltInStyleIndex index = GetBuiltInStyleIndex();
        if (index.IdByBuiltInId.TryGetValue(builtInStyleId, out string? styleId)) return styleId;
        // A definition under the English identifier counts only when it has the built-in style's type.
        return index.Definitions.TryGetValue(builtInStyleId, out var definition) &&
            WordBuiltInStyleNames.Classify(builtInStyleId, null, definition.Type) != null ? builtInStyleId : null;
    }

    /// <summary>
    /// Returns the identifier to apply for a built-in style: the document's own style for the built-in name, or the
    /// English identifier, which <see cref="EnsureStyleDefinition"/> creates when the document has no such style.
    /// </summary>
    internal string ResolveBuiltInStyleId(string builtInStyleId) => FindBuiltInStyleId(builtInStyleId) ?? builtInStyleId;

    /// <summary>Applies a built-in style, reusing the document's localized definition when it has one.</summary>
    /// <returns>The identifier the content now references.</returns>
    internal string UseBuiltInStyle(string builtInStyleId) {
        string? styleId = FindBuiltInStyleId(builtInStyleId);
        if (styleId != null && !string.Equals(styleId, builtInStyleId, StringComparison.Ordinal)) return styleId;
        EnsureStyleDefinition(builtInStyleId);
        return builtInStyleId;
    }

    /// <summary>
    /// Returns the identifier to store for a requested style identifier. An English built-in identifier the document
    /// does not define maps to the document's localized definition of that built-in style; any other identifier is kept
    /// and its library definition is added when the document lacks it.
    /// </summary>
    internal string UseStyleId(string styleId) {
        if (WordBuiltInStyleNames.IsBuiltInId(styleId) && GetUnparsedDefaultPackageStyles() == null &&
            !GetBuiltInStyleIndex().Definitions.ContainsKey(styleId)) {
            return UseBuiltInStyle(styleId);
        }
        EnsureStyleDefinition(styleId);
        return styleId;
    }

    /// <summary>
    /// Classifies a style identifier this document uses: the English identifier of the built-in style it is, found
    /// through the definition's built-in name, or the identifier itself when it is an English built-in identifier
    /// without a definition naming another style.
    /// </summary>
    /// <returns>The English built-in identifier, or <see langword="null"/> for a custom style.</returns>
    internal string? GetBuiltInStyleId(string? styleId) {
        if (string.IsNullOrEmpty(styleId)) return null;
        if (GetUnparsedDefaultPackageStyles() != null) {
            return WordBuiltInStyleNames.IsBuiltInId(styleId) ? styleId : null;
        }
        BuiltInStyleIndex index = GetBuiltInStyleIndex();
        if (index.Definitions.TryGetValue(styleId!, out var definition)) {
            return WordBuiltInStyleNames.Classify(styleId, definition.Name, definition.Type);
        }
        return WordBuiltInStyleNames.IsBuiltInId(styleId) ? styleId : null;
    }

    /// <summary>
    /// Returns the left indentation, in twips, a paragraph style gives through its <c>w:basedOn</c> chain and then the
    /// document's default paragraph properties, or 0.
    /// </summary>
    internal int GetStyleIndentationBeforeTwips(string styleId) {
        WordStyleCatalog catalog = WordStyleCatalog.For(this);
        foreach (Style style in catalog.EnumerateChain(styleId, StyleValues.Paragraph)) {
            if (TryReadLeftIndent(style.StyleParagraphProperties?.Indentation, out int twips)) return twips;
        }
        return TryReadLeftIndent(catalog.DocumentDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.Indentation, out int defaultTwips)
            ? defaultTwips
            : 0;

        static bool TryReadLeftIndent(Indentation? indentation, out int twips) {
            twips = 0;
            string? left = indentation?.Left?.Value ?? indentation?.Start?.Value;
            return left != null && int.TryParse(left, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out twips);
        }
    }

    /// <summary>Returns the heading level, 1 to 9, of a style identifier, through its built-in name; otherwise 0.</summary>
    internal int GetHeadingLevel(string? styleId) => WordBuiltInStyleNames.GetHeadingLevel(GetBuiltInStyleId(styleId));

    /// <summary>
    /// Points generated content that references built-in styles by English identifier at the document's own
    /// definitions of those styles, so a localized template keeps one definition per built-in style.
    /// </summary>
    internal void LocalizeBuiltInStyleReferences(DocumentFormat.OpenXml.OpenXmlElement root) {
        if (GetUnparsedDefaultPackageStyles() != null) return;
        BuiltInStyleIndex index = GetBuiltInStyleIndex();
        foreach (ParagraphStyleId reference in root.Descendants<ParagraphStyleId>()) {
            // A reference to an identifier the document defines, such as an English Quote, is kept as written.
            if (reference.Val?.Value is { } styleId && WordBuiltInStyleNames.IsBuiltInId(styleId) &&
                !index.Definitions.ContainsKey(styleId) &&
                FindBuiltInStyleId(styleId) is { } localized && !string.Equals(localized, styleId, StringComparison.Ordinal)) {
                reference.Val = localized;
            }
        }
    }

    /// <summary>Drops the built-in style index after a style is added, removed or renamed.</summary>
    internal void InvalidateBuiltInStyleIndex() => _builtInStyleIndex = null;

    private BuiltInStyleIndex GetBuiltInStyleIndex() {
        StyleDefinitionsPart? part = _wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart;
        BuiltInStyleIndex? index = _builtInStyleIndex;
        // Reading a document classifies styles from the part's XML, so inspection does not load the styles DOM.
        if (part != null && !part.IsRootElementLoaded) {
            if (index != null && ReferenceEquals(index.Source, part)) return index;
            if (WordStyleDefinitionsReader.TryRead(part) is { } streamed) {
                index = new BuiltInStyleIndex(part, streamed.Styles.Select(style => (style.StyleId, style.Name, style.Type)));
                _builtInStyleIndex = index;
                return index;
            }
        }
        Styles? styles = part?.Styles;
        // Adding, removing and renaming styles through OfficeIMO invalidates the index; a replaced part rebuilds it.
        if (index != null && ReferenceEquals(index.Source, styles)) return index;
        index = new BuiltInStyleIndex(styles, styles?.Elements<Style>().Select(style =>
            (style.StyleId?.Value, style.StyleName?.Val?.Value, WordBuiltInStyleNames.StyleType(style))));
        _builtInStyleIndex = index;
        return index;
    }

    private sealed class BuiltInStyleIndex {
        internal BuiltInStyleIndex(object? source, IEnumerable<(string? Id, string? Name, string? Type)>? styles) {
            Source = source;
            if (styles == null) return;
            foreach ((string? id, string? name, string? type) in styles) {
                if (string.IsNullOrEmpty(id) || Definitions.ContainsKey(id!)) continue;
                Definitions.Add(id!, (name, type));
                // Word itself keeps the first definition of a name; a style of another type never stands in.
                if (WordBuiltInStyleNames.GetId(name) is { } builtInId &&
                    WordBuiltInStyleNames.Classify(id, name, type) == builtInId && !IdByBuiltInId.ContainsKey(builtInId)) {
                    IdByBuiltInId.Add(builtInId, id!);
                }
            }
        }

        /// <summary>The style part read as XML, or the loaded styles DOM.</summary>
        internal object? Source { get; }
        internal Dictionary<string, (string? Name, string? Type)> Definitions { get; } = new(StringComparer.Ordinal);
        internal Dictionary<string, string> IdByBuiltInId { get; } = new(StringComparer.Ordinal);
    }
}

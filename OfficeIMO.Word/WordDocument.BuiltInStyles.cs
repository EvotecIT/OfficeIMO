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
        BuiltInStyleLookup lookup = GetBuiltInStyleLookup();
        if (lookup.TryGetStyleIdForBuiltInName(builtInStyleId, out string? styleId)) return styleId;
        // A definition under the English identifier counts only when it has the built-in style's type.
        return lookup.TryGetDefinition(builtInStyleId, out _, out string? type) &&
            WordBuiltInStyleNames.Classify(builtInStyleId, null, type) != null ? builtInStyleId : null;
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
            !GetBuiltInStyleLookup().TryGetDefinition(styleId, out _, out _)) {
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
        if (GetBuiltInStyleLookup().TryGetDefinition(styleId!, out string? name, out string? type)) {
            return WordBuiltInStyleNames.Classify(styleId, name, type);
        }
        return WordBuiltInStyleNames.IsBuiltInId(styleId) ? styleId : null;
    }

    /// <summary>
    /// Returns the left indentation, in twips, a paragraph style gives through its <c>w:basedOn</c> chain and then the
    /// document's default paragraph properties, or 0. No identifier, or one the document does not define, means the
    /// default paragraph style, as Word reads it.
    /// </summary>
    internal int GetStyleIndentationBeforeTwips(string? styleId) {
        WordStyleCatalog catalog = WordStyleCatalog.For(this);
        IEnumerable<Style> chain = catalog.EnumerateChain(styleId, StyleValues.Paragraph);
        if (!string.IsNullOrEmpty(styleId) && !chain.Any()) chain = catalog.EnumerateChain(null, StyleValues.Paragraph);
        foreach (Style style in chain) {
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
        BuiltInStyleLookup lookup = GetBuiltInStyleLookup();
        foreach (ParagraphStyleId reference in root.Descendants<ParagraphStyleId>()) {
            // A reference to an identifier the document defines, such as an English Quote, is kept as written.
            if (reference.Val?.Value is { } styleId && WordBuiltInStyleNames.IsBuiltInId(styleId) &&
                !lookup.TryGetDefinition(styleId, out _, out _) &&
                FindBuiltInStyleId(styleId) is { } localized && !string.Equals(localized, styleId, StringComparison.Ordinal)) {
                reference.Val = localized;
            }
        }
    }

    /// <summary>
    /// Returns the left indentation, in twips, a paragraph shows before numbering: its direct indent, otherwise the one its
    /// paragraph style gives. Indents added for quotes, lists and containers go on top of this, because a direct indent
    /// replaces the style's. Numbered paragraphs take their indent from the list, so they start from their direct indent.
    /// </summary>
    internal int GetEffectiveIndentationBeforeTwips(WordParagraph paragraph) {
        if (paragraph.IndentationBefore is { } direct) return checked((int)direct.ToTwips());
        if (paragraph._paragraph.ParagraphProperties?.NumberingProperties != null) return 0;
        return GetStyleIndentationBeforeTwips(paragraph.StyleId);
    }

    /// <summary>Drops the built-in style index after a style is added, removed or renamed.</summary>
    internal void InvalidateBuiltInStyleIndex() => _builtInStyleIndex = null;

    private BuiltInStyleLookup GetBuiltInStyleLookup() {
        StyleDefinitionsPart? part = _wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart;
        // While the styles part is unloaded, the shared style catalog answers from its identifier and name indexes,
        // so reading and inspecting a document never loads the styles DOM.
        if (part != null && !part.IsRootElementLoaded) {
            WordStyleCatalog catalog = WordStyleCatalog.For(this);
            if (catalog.IsStreamed) return new CatalogLookup(catalog);
        }
        // Once the raw package is handed out, edits to the styles DOM can't be observed, so the index is not reused.
        BuiltInStyleIndex? index = _openXmlDocumentAccessed ? null : _builtInStyleIndex;
        Styles? styles = part?.Styles;
        // Over the loaded DOM the catalog scans every style per lookup, and paragraph styles are classified for every
        // paragraph, so a small index is kept. Adding, removing and renaming styles through OfficeIMO invalidates it,
        // and a replaced part rebuilds it.
        if (index != null && ReferenceEquals(index.Styles, styles)) return index;
        index = new BuiltInStyleIndex(styles);
        _builtInStyleIndex = index;
        return index;
    }

    /// <summary>The style facts built-in style resolution needs: definitions by identifier and by built-in name.</summary>
    private abstract class BuiltInStyleLookup {
        /// <summary>The name and <c>w:type</c> of the first definition with this identifier.</summary>
        internal abstract bool TryGetDefinition(string styleId, out string? name, out string? type);

        /// <summary>The identifier of the first definition named after this built-in style, with the built-in style's type.</summary>
        internal abstract bool TryGetStyleIdForBuiltInName(string builtInStyleId, out string? styleId);
    }

    private sealed class CatalogLookup : BuiltInStyleLookup {
        private readonly WordStyleCatalog _catalog;

        internal CatalogLookup(WordStyleCatalog catalog) => _catalog = catalog;

        internal override bool TryGetDefinition(string styleId, out string? name, out string? type) {
            WordStyleDefinitionsReader.StyleDefinition? definition = _catalog.FindDefinitionById(styleId);
            name = definition?.Name;
            type = definition?.Type;
            return definition != null;
        }

        internal override bool TryGetStyleIdForBuiltInName(string builtInStyleId, out string? styleId) {
            styleId = null;
            if (!WordBuiltInStyleNames.TryGetName(builtInStyleId, out string name)) return false;
            WordStyleDefinitionsReader.StyleDefinition? definition = _catalog.FindDefinitionByName(name);
            if (definition?.StyleId is not { Length: > 0 } id ||
                WordBuiltInStyleNames.Classify(id, definition.Name, definition.Type) != builtInStyleId) return false;
            styleId = id;
            return true;
        }
    }

    private sealed class BuiltInStyleIndex : BuiltInStyleLookup {
        private readonly Dictionary<string, (string? Name, string? Type)> _definitions = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _idByBuiltInId = new(StringComparer.Ordinal);

        internal BuiltInStyleIndex(Styles? styles) {
            Styles = styles;
            if (styles == null) return;
            foreach (Style style in styles.Elements<Style>()) {
                string? id = style.StyleId?.Value;
                if (string.IsNullOrEmpty(id) || _definitions.ContainsKey(id!)) continue;
                string? name = style.StyleName?.Val?.Value;
                string? type = WordBuiltInStyleNames.StyleType(style);
                _definitions.Add(id!, (name, type));
                // Word itself keeps the first definition of a name; a style of another type never stands in.
                if (WordBuiltInStyleNames.GetId(name) is { } builtInId &&
                    WordBuiltInStyleNames.Classify(id, name, type) == builtInId && !_idByBuiltInId.ContainsKey(builtInId)) {
                    _idByBuiltInId.Add(builtInId, id!);
                }
            }
        }

        /// <summary>The loaded styles DOM the index was built from.</summary>
        internal Styles? Styles { get; }

        internal override bool TryGetDefinition(string styleId, out string? name, out string? type) {
            bool found = _definitions.TryGetValue(styleId, out (string? Name, string? Type) definition);
            name = definition.Name;
            type = definition.Type;
            return found;
        }

        internal override bool TryGetStyleIdForBuiltInName(string builtInStyleId, out string? styleId) {
            bool found = _idByBuiltInId.TryGetValue(builtInStyleId, out string? id);
            styleId = id;
            return found;
        }
    }
}

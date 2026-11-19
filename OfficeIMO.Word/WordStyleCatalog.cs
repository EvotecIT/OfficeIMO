using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using StyleDefinition = OfficeIMO.Word.WordStyleDefinitionsReader.StyleDefinition;

namespace OfficeIMO.Word;

/// <summary>
/// The style definitions of one document as read-only callers see them: the styles in document order, lookup by
/// identifier, default styles, <c>w:basedOn</c> chains and the document defaults.
/// </summary>
/// <remarks>
/// While the styles part is unloaded the catalog is read from its stream once and kept for the document until
/// the part changes. Each style's facts are available at once; its properties are parsed into a detached
/// <see cref="Style"/> only when a caller asks for them, so a reader pays only for the styles its content uses.
/// When the styles part is already loaded (it may carry unsaved edits), when the package was opened with
/// markup-compatibility processing, or when the raw Open XML document has been handed out, the catalog answers
/// from the loaded <see cref="Styles"/> element instead, exactly as the DOM reads it.
/// Elements returned from a streamed catalog are detached copies and must not be modified.
/// </remarks>
internal sealed class WordStyleCatalog {
    private static readonly WordStyleCatalog Empty = new(null, null, null, 0);

    private readonly Styles? _loaded;
    private readonly WordStyleDefinitionsReader.StyleDefinitionSet? _streamed;
    private readonly Dictionary<string, StyleDefinition>? _firstById;
    private readonly StylesPart? _part;
    private readonly int _revision;

    private WordStyleCatalog(Styles? loaded, WordStyleDefinitionsReader.StyleDefinitionSet? streamed, StylesPart? part, int revision) {
        _loaded = loaded;
        _streamed = streamed;
        _part = part;
        _revision = revision;
        if (streamed == null) return;
        _firstById = new Dictionary<string, StyleDefinition>(StringComparer.Ordinal);
        foreach (StyleDefinition definition in streamed.Styles) {
            if (definition.StyleId != null && !_firstById.ContainsKey(definition.StyleId)) _firstById.Add(definition.StyleId, definition);
        }
    }

    /// <summary>Returns the catalog of <paramref name="document"/>'s styles part, reusing the one already read.</summary>
    internal static WordStyleCatalog For(WordDocument document) {
        StylesPart? part = document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart;
        if (part == null) return Empty;
        int revision = Volatile.Read(ref document._styleDefinitionsRevision);
        WordStyleCatalog? cached = Volatile.Read(ref document._styleCatalog);
        if (!part.IsRootElementLoaded && !document._openXmlDocumentAccessed) {
            if (cached != null && cached._streamed != null && ReferenceEquals(cached._part, part) && cached._revision == revision) return cached;
            WordStyleDefinitionsReader.StyleDefinitionSet? streamed = WordStyleDefinitionsReader.TryRead(part);
            if (streamed != null) {
                var created = new WordStyleCatalog(null, streamed, part, revision);
                Volatile.Write(ref document._styleCatalog, created);
                return created;
            }
        }
        Styles? styles = part.Styles;
        if (cached != null && cached._streamed == null && ReferenceEquals(cached._loaded, styles)) return cached;
        var loaded = styles == null ? Empty : new WordStyleCatalog(styles, null, part, revision);
        Volatile.Write(ref document._styleCatalog, loaded);
        return loaded;
    }

    /// <summary>A catalog over a <see cref="Styles"/> element the caller already holds, such as one being edited.</summary>
    internal static WordStyleCatalog ForLoaded(Styles? styles) => styles == null ? Empty : new WordStyleCatalog(styles, null, null, 0);

    /// <summary>Whether the catalog was read from the unloaded styles part rather than its loaded DOM.</summary>
    internal bool IsStreamed => _streamed != null;

    /// <summary>The streamed definitions, or null when the catalog reads the loaded DOM.</summary>
    internal WordStyleDefinitionsReader.StyleDefinitionSet? Streamed => _streamed;

    /// <summary>Every top-level style in document order. Over the DOM, each call reads the current elements.</summary>
    internal IReadOnlyList<StyleDefinition> Definitions {
        get {
            if (_streamed != null) return _streamed.Styles;
            if (_loaded == null) return Array.Empty<StyleDefinition>();
            var definitions = new List<StyleDefinition>();
            foreach (Style style in _loaded.Elements<Style>()) definitions.Add(new StyleDefinition(style));
            return definitions;
        }
    }

    /// <summary>The number of top-level styles.</summary>
    internal int Count => _streamed != null ? _streamed.Styles.Count : _loaded?.Elements<Style>().Count() ?? 0;

    /// <summary>Whether any style has <paramref name="styleId"/>, whatever its type.</summary>
    internal bool Contains(string styleId) => _firstById != null
        ? _firstById.ContainsKey(styleId)
        : _loaded?.Elements<Style>().Any(style => string.Equals(style.StyleId?.Value, styleId, StringComparison.Ordinal)) == true;

    /// <summary>The document defaults (<c>w:docDefaults</c>), or null when the styles part declares none.</summary>
    internal DocDefaults? DocumentDefaults => _streamed != null ? _streamed.DocumentDefaults : _loaded?.DocDefaults;

    /// <summary>The identifier of the first style of <paramref name="type"/> marked as the default.</summary>
    internal string? DefaultStyleId(StyleValues type) {
        if (_streamed != null) {
            string name = TypeName(type);
            foreach (StyleDefinition definition in _streamed.Styles) {
                if (definition.IsDefault && string.Equals(definition.Type, name, StringComparison.Ordinal)) return definition.StyleId;
            }
            return null;
        }
        return _loaded?.Elements<Style>().FirstOrDefault(style => style.Type?.Value == type && style.Default?.Value == true)?.StyleId?.Value;
    }

    /// <summary>The first style with <paramref name="styleId"/> that has no type or <paramref name="type"/>.</summary>
    internal Style? Find(string? styleId, StyleValues type) => FindDefinition(styleId, type)?.Element;

    /// <summary>The first style with <paramref name="styleId"/>, whatever its type.</summary>
    internal Style? FindById(string? styleId) {
        if (styleId == null) return null;
        if (_firstById != null) return _firstById.TryGetValue(styleId!, out StyleDefinition? definition) ? definition.Element : null;
        return _loaded?.Elements<Style>().FirstOrDefault(style => string.Equals(style.StyleId?.Value, styleId, StringComparison.Ordinal));
    }

    /// <summary>The facts of the first style with <paramref name="styleId"/> that has no type or <paramref name="type"/>.</summary>
    internal StyleDefinition? FindDefinition(string? styleId, StyleValues type) {
        if (string.IsNullOrEmpty(styleId)) return null;
        if (_streamed != null) {
            string name = TypeName(type);
            if (!_firstById!.TryGetValue(styleId!, out StyleDefinition? first)) return null;
            if (first.Matches(name)) return first;
            // A later definition with the same identifier can still match the requested type.
            foreach (StyleDefinition definition in _streamed.Styles) {
                if (string.Equals(definition.StyleId, styleId, StringComparison.Ordinal) && definition.Matches(name)) return definition;
            }
            return null;
        }
        Style? style = _loaded?.Elements<Style>().FirstOrDefault(candidate =>
            string.Equals(candidate.StyleId?.Value, styleId, StringComparison.Ordinal) &&
            (candidate.Type == null || !candidate.Type.HasValue || candidate.Type.Value == type));
        return style == null ? null : new StyleDefinition(style);
    }

    /// <summary>
    /// Enumerates a style and its ancestors in override order, using the default style of the type only when no
    /// identifier is supplied. Missing styles, mismatched types and inheritance cycles end the chain.
    /// </summary>
    internal IEnumerable<Style> EnumerateChain(string? styleId, StyleValues type) {
        if (_streamed == null && _loaded == null) yield break;
        if (string.IsNullOrWhiteSpace(styleId)) styleId = DefaultStyleId(type);
        if (string.IsNullOrWhiteSpace(styleId)) yield break;

        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
            if (_streamed != null) {
                StyleDefinition? definition = FindDefinition(styleId, type);
                if (definition == null) yield break;
                yield return definition.Element;
                styleId = definition.BasedOn;
            } else {
                Style? style = _loaded!.Elements<Style>().FirstOrDefault(candidate =>
                    string.Equals(candidate.StyleId?.Value, styleId, StringComparison.Ordinal) &&
                    (candidate.Type == null || !candidate.Type.HasValue || candidate.Type.Value == type));
                if (style == null) yield break;
                yield return style;
                styleId = style.BasedOn?.Val?.Value;
            }
        }
    }

    private static string TypeName(StyleValues type) {
        if (type == StyleValues.Paragraph) return "paragraph";
        if (type == StyleValues.Character) return "character";
        if (type == StyleValues.Table) return "table";
        if (type == StyleValues.Numbering) return "numbering";
        return type.ToString();
    }
}

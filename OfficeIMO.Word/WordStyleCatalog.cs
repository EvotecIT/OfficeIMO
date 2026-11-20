using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using StyleDefinition = OfficeIMO.Word.WordStyleDefinitionsReader.StyleDefinition;

namespace OfficeIMO.Word;

/// <summary>
/// The style definitions of one styles part as read-only callers see them: the styles in document order, lookup by
/// identifier, default styles, <c>w:basedOn</c> chains and the document defaults.
/// </summary>
/// <remarks>
/// While the styles part is unloaded the catalog is read from its stream once and kept with the part until the
/// part is loaded or OfficeIMO replaces its content. Each style's facts are available at once; its properties are parsed into a detached
/// <see cref="Style"/> only when a caller asks for them, so a reader pays only for the styles its content uses.
/// When the styles part is already loaded (it may carry unsaved edits), when the package was opened with
/// markup-compatibility processing, or when the raw Open XML document has been handed out, the catalog answers
/// from the loaded <see cref="Styles"/> element instead, exactly as the DOM reads it.
/// Elements returned from a streamed catalog are detached copies and must not be modified.
/// </remarks>
internal sealed class WordStyleCatalog {
    private static readonly WordStyleCatalog Empty = new(null, null);
    private static readonly ConditionalWeakTable<StylesPart, PartState> Parts = new();

    private readonly Styles? _loaded;
    private readonly WordStyleDefinitionsReader.StyleDefinitionSet? _streamed;
    private readonly Dictionary<string, StyleDefinition>? _firstById;
    private readonly Dictionary<string, StyleDefinition>? _firstByName;

    /// <summary>The catalog last read for one part.</summary>
    private sealed class PartState {
        internal WordStyleCatalog? Catalog;
    }

    private WordStyleCatalog(Styles? loaded, WordStyleDefinitionsReader.StyleDefinitionSet? streamed) {
        _loaded = loaded;
        _streamed = streamed;
        if (streamed == null) return;
        _firstById = new Dictionary<string, StyleDefinition>(StringComparer.Ordinal);
        foreach (StyleDefinition definition in streamed.Styles) {
            if (definition.StyleId != null && !_firstById.ContainsKey(definition.StyleId)) _firstById.Add(definition.StyleId, definition);
            string? name = definition.Name?.Trim();
            if (name != null) {
                _firstByName ??= new Dictionary<string, StyleDefinition>(StringComparer.OrdinalIgnoreCase);
                if (!_firstByName.ContainsKey(name)) _firstByName.Add(name, definition);
            }
        }
    }

    /// <summary>Returns the catalog of <paramref name="document"/>'s styles part, reusing the one already read.</summary>
    internal static WordStyleCatalog For(WordDocument document) {
        // Once the raw package has been handed out, its styles parts, including ones added later, are read from their DOM.
        return For(document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart);
    }

    /// <summary>Returns the catalog of the styles part of the package that owns <paramref name="mainPart"/>.</summary>
    internal static WordStyleCatalog For(MainDocumentPart? mainPart) => For(mainPart?.StyleDefinitionsPart);

    /// <summary>
    /// Returns the catalog of <paramref name="part"/>, which may be the styles or the styles-with-effects part,
    /// reusing the one already read while the part stays unloaded.
    /// </summary>
    internal static WordStyleCatalog For(StylesPart? part) {
        if (part == null) return Empty;
        PartState state = Parts.GetValue(part, _ => new PartState());
        WordStyleCatalog? cached = Volatile.Read(ref state.Catalog);
        if (WordStyleDefinitionsReader.CanStream(part)) {
            if (cached?._streamed != null) return cached;
            WordStyleDefinitionsReader.StyleDefinitionSet? streamed = WordStyleDefinitionsReader.TryRead(part);
            if (streamed != null) {
                var created = new WordStyleCatalog(null, streamed);
                Volatile.Write(ref state.Catalog, created);
                return created;
            }
        }
        Styles? styles = part.Styles;
        if (cached != null && cached._streamed == null && ReferenceEquals(cached._loaded, styles)) return cached;
        var loaded = styles == null ? Empty : new WordStyleCatalog(styles, null);
        Volatile.Write(ref state.Catalog, loaded);
        return loaded;
    }

    /// <summary>Forgets the catalog read from <paramref name="part"/> after OfficeIMO replaced the part's content.</summary>
    internal static void Invalidate(StylesPart? part) {
        if (part != null && Parts.TryGetValue(part, out PartState? state)) Volatile.Write(ref state.Catalog, null);
    }

    /// <summary>A catalog over a <see cref="Styles"/> element the caller already holds, such as one being edited.</summary>
    internal static WordStyleCatalog ForLoaded(Styles? styles) => styles == null ? Empty : new WordStyleCatalog(styles, null);

    /// <summary>Whether the catalog was read from the unloaded styles part rather than its loaded DOM.</summary>
    internal bool IsStreamed => _streamed != null;

    /// <summary>The loaded <see cref="Styles"/> element the catalog reads, or null when it was streamed.</summary>
    internal Styles? Loaded => _loaded;

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

    /// <summary>The facts of the first style with <paramref name="styleId"/>, whatever its type.</summary>
    internal StyleDefinition? FindDefinitionById(string? styleId) {
        if (string.IsNullOrEmpty(styleId)) return null;
        if (_firstById != null) return _firstById.TryGetValue(styleId!, out StyleDefinition? definition) ? definition : null;
        Style? style = _loaded?.Elements<Style>().FirstOrDefault(candidate => string.Equals(candidate.StyleId?.Value, styleId, StringComparison.Ordinal));
        return style == null ? null : new StyleDefinition(style);
    }

    /// <summary>
    /// The facts of the first style whose <c>w:name</c> matches <paramref name="name"/>, both trimmed and ignoring case.
    /// Built-in style identifiers are localized in some documents while their names stay English.
    /// </summary>
    internal StyleDefinition? FindDefinitionByName(string name) {
        string trimmed = name.Trim();
        if (_streamed != null) return _firstByName != null && _firstByName.TryGetValue(trimmed, out StyleDefinition? definition) ? definition : null;
        Style? style = _loaded?.Elements<Style>().FirstOrDefault(candidate =>
            string.Equals(candidate.StyleName?.Val?.Value?.Trim(), trimmed, StringComparison.OrdinalIgnoreCase));
        return style == null ? null : new StyleDefinition(style);
    }

    /// <summary>
    /// The first style named <paramref name="name"/> that content can reference and <paramref name="accept"/> takes:
    /// it has an identifier, and that identifier resolves back to this definition rather than to an earlier style that
    /// holds the same identifier. Every name-based lookup goes through this filter, so malformed files with shared names,
    /// shared identifiers or missing identifiers resolve the same way while the styles part is unloaded or loaded.
    /// </summary>
    internal StyleDefinition? FindReferenceableByName(string name, Func<StyleDefinition, bool> accept) {
        foreach (StyleDefinition definition in FindDefinitionsByName(name)) {
            if (string.IsNullOrEmpty(definition.StyleId) || !accept(definition)) continue;
            if (_firstById != null) {
                if (_firstById.TryGetValue(definition.StyleId!, out StyleDefinition? first) && ReferenceEquals(first, definition)) return definition;
                continue;
            }
            Style? firstWithId = _loaded?.Elements<Style>().FirstOrDefault(style => string.Equals(style.StyleId?.Value, definition.StyleId, StringComparison.Ordinal));
            if (firstWithId != null && ReferenceEquals(firstWithId, definition.Element)) return definition;
        }
        return null;
    }

    /// <summary>
    /// The facts of every style whose <c>w:name</c> matches <paramref name="name"/>, in document order, so a caller
    /// that needs one type can take the first of that type even when a style of another type shares the name.
    /// </summary>
    internal IEnumerable<StyleDefinition> FindDefinitionsByName(string name) {
        string trimmed = name.Trim();
        if (_streamed != null) {
            if (_firstByName == null || !_firstByName.TryGetValue(trimmed, out StyleDefinition? first)) yield break;
            yield return first;
            // Later styles with the same name are rare, so only a caller that skips the first one pays for the scan.
            bool afterFirst = false;
            foreach (StyleDefinition definition in _streamed.Styles) {
                if (ReferenceEquals(definition, first)) { afterFirst = true; continue; }
                if (afterFirst && string.Equals(definition.Name?.Trim(), trimmed, StringComparison.OrdinalIgnoreCase)) yield return definition;
            }
            yield break;
        }
        if (_loaded == null) yield break;
        foreach (Style style in _loaded.Elements<Style>()) {
            if (string.Equals(style.StyleName?.Val?.Value?.Trim(), trimmed, StringComparison.OrdinalIgnoreCase)) yield return new StyleDefinition(style);
        }
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

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Word;

/// <summary>
/// Resolves the effective numbering attached to a paragraph, including numbering inherited
/// through paragraph styles and direct <c>w:numId="0"</c> cancellation.
/// </summary>
internal static class WordListNumberingResolver {
    private const int MaximumStyleInheritanceDepth = 256;
    private static readonly ConditionalWeakTable<MainDocumentPart, CachedStyleCatalog> StyleCatalogs = new();
    private static readonly ConditionalWeakTable<Numbering, NumberingRevision> NumberingRevisions = new();

    /// <summary>
    /// Resolves abstract IDs to the first definition with the same authored numbering identity.
    /// Definitions without an identity remain independent; source XML is preserved.
    /// </summary>
    internal static Dictionary<int, AbstractNum> GetCanonicalAbstractDefinitions(Numbering numbering) {
        var definitions = new Dictionary<int, AbstractNum>();
        var identities = new Dictionary<string, AbstractNum>(StringComparer.OrdinalIgnoreCase);
        foreach (AbstractNum definition in numbering.Elements<AbstractNum>()) {
            if (definition.AbstractNumberId?.Value is not int id || definitions.ContainsKey(id)) continue;
            AbstractNum canonical = definition;
            string? identity = definition.Nsid?.Val?.Value;
            if (!string.IsNullOrEmpty(identity)) {
                if (identities.TryGetValue(identity!, out AbstractNum? first)) canonical = first;
                else identities.Add(identity!, definition);
            }
            definitions.Add(id, canonical);
        }
        return definitions;
    }

    private sealed class NumberingRevision {
        internal int Value;
    }

    private sealed class CachedStyleCatalog {
        internal StylesPart? StylesPart;
        internal StylesPart? StylesWithEffectsPart;
        internal Styles? Styles;
        internal Styles? StylesWithEffects;
        internal Numbering? Numbering;
        internal int StyleCount;
        internal int EffectStyleCount;
        internal int NumberingCount;
        internal int NumberingRevision;
        internal StyleCatalog? Catalog;
    }

    internal sealed class StyleCatalog {
        internal StyleCatalog(MainDocumentPart? mainPart, bool memoizeResolutions = false) {
            if (memoizeResolutions) {
                NumberingByStyle = new Dictionary<string, NumberingProperties?>(StringComparer.Ordinal);
                LinkedStyleByStyle = new Dictionary<string, ResolvedNumbering?>(StringComparer.Ordinal);
                LinkedLevelByStyle = new Dictionary<(int NumberId, string StyleId), int?>();
            }
            ById = new Dictionary<string, ParagraphStyleEntry>(StringComparer.Ordinal);
            string? defaultStyleId = null;
            StreamedStyleDefinitions = WordStyleDefinitionsReader.TryRead(mainPart?.StyleDefinitionsPart);
            AddParagraphStyles(mainPart?.StyleDefinitionsPart, StreamedStyleDefinitions, ref defaultStyleId);
            AddParagraphStyles(mainPart?.StylesWithEffectsPart, WordStyleDefinitionsReader.TryRead(mainPart?.StylesWithEffectsPart), ref defaultStyleId);
            DefaultStyleId = defaultStyleId;
            LinkedLevels = new Dictionary<(int NumberId, string StyleId), int>();
            LinkedStyles = new Dictionary<string, ResolvedNumbering>(StringComparer.Ordinal);
            ListDefinitions = WordDocumentTraversal.BuildListNumberingDefinitions(mainPart);
            Numbering? numbering = mainPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) return;
            Dictionary<int, AbstractNum> abstracts = GetCanonicalAbstractDefinitions(numbering);
            foreach (NumberingInstance instance in numbering.Elements<NumberingInstance>()) {
                if (instance.NumberID?.Value is not int numberId ||
                    instance.AbstractNumId?.Val?.Value is not int abstractId ||
                    !abstracts.TryGetValue(abstractId, out AbstractNum? abstractNum)) continue;
                Dictionary<int, Level> overrides = instance.Elements<LevelOverride>()
                    .Where(levelOverride => levelOverride.LevelIndex?.Value != null && levelOverride.GetFirstChild<Level>() != null)
                    .GroupBy(levelOverride => levelOverride.LevelIndex!.Value)
                    .ToDictionary(group => group.Key, group => group.First().GetFirstChild<Level>()!);
                foreach (Level abstractLevel in abstractNum.Elements<Level>()) {
                    if (abstractLevel.LevelIndex?.Value is not int index) continue;
                    Level level = overrides.TryGetValue(index, out Level? replacement) ? replacement : abstractLevel;
                    if (level.GetFirstChild<ParagraphStyleIdInLevel>()?.Val?.Value is string linkedStyle) {
                        LinkedLevels[(numberId, linkedStyle)] = index;
                        if (!LinkedStyles.ContainsKey(linkedStyle)) LinkedStyles.Add(linkedStyle, new ResolvedNumbering(numberId, index));
                    }
                }
                foreach (KeyValuePair<int, Level> levelOverride in overrides) {
                    if (levelOverride.Value.GetFirstChild<ParagraphStyleIdInLevel>()?.Val?.Value is string linkedStyle) {
                        LinkedLevels[(numberId, linkedStyle)] = levelOverride.Key;
                        if (!LinkedStyles.ContainsKey(linkedStyle)) LinkedStyles.Add(linkedStyle, new ResolvedNumbering(numberId, levelOverride.Key));
                    }
                }
            }
        }

        private void AddParagraphStyles(StylesPart? part, WordStyleDefinitionsReader.StyleDefinitionSet? streamed,
            ref string? defaultStyleId) {
            if (part == null) return;
            if (streamed != null) {
                foreach (WordStyleDefinitionsReader.StyleDefinition definition in streamed.Styles) {
                    if (!definition.IsParagraphStyle || string.IsNullOrWhiteSpace(definition.StyleId) || ById.ContainsKey(definition.StyleId!)) continue;
                    ById.Add(definition.StyleId!, new ParagraphStyleEntry(definition));
                    if (defaultStyleId == null && definition.IsDefault) defaultStyleId = definition.StyleId;
                }
                return;
            }
            Styles? styles = part.Styles;
            if (styles == null) return;
            foreach (Style style in styles.Elements<Style>()) {
                string? styleId = style.StyleId?.Value;
                if (style.Type?.Value != StyleValues.Paragraph || string.IsNullOrWhiteSpace(styleId) || ById.ContainsKey(styleId!)) continue;
                ById.Add(styleId!, new ParagraphStyleEntry(style));
                if (defaultStyleId == null && style.Default?.Value == true) defaultStyleId = styleId;
            }
        }

        /// <summary>Paragraph styles by identifier, first definition winning, from the styles part and then the styles-with-effects part.</summary>
        internal Dictionary<string, ParagraphStyleEntry> ById { get; }

        /// <summary>
        /// Style definitions read from the unloaded styles part without building its DOM, or null when the catalog
        /// read the loaded <see cref="Styles"/> element instead.
        /// </summary>
        internal WordStyleDefinitionsReader.StyleDefinitionSet? StreamedStyleDefinitions { get; }
        internal string? DefaultStyleId { get; }
        internal Dictionary<(int NumberId, string StyleId), int> LinkedLevels { get; }
        internal Dictionary<string, ResolvedNumbering> LinkedStyles { get; }
        internal IReadOnlyDictionary<int, WordDocumentTraversal.ListNumberingDefinition> ListDefinitions { get; }
        internal Dictionary<string, NumberingProperties?>? NumberingByStyle { get; }
        internal Dictionary<string, ResolvedNumbering?>? LinkedStyleByStyle { get; }
        internal Dictionary<(int NumberId, string StyleId), int?>? LinkedLevelByStyle { get; }
    }

    /// <summary>
    /// A paragraph style's base style and numbering, read live from a loaded <see cref="Style"/> element or from a
    /// definition streamed out of an unloaded styles part.
    /// </summary>
    internal sealed class ParagraphStyleEntry {
        private readonly Style? _style;
        private readonly WordStyleDefinitionsReader.StyleDefinition? _definition;

        internal ParagraphStyleEntry(Style style) => _style = style;

        internal ParagraphStyleEntry(WordStyleDefinitionsReader.StyleDefinition definition) => _definition = definition;

        internal string? BasedOn => _style != null ? _style.BasedOn?.Val?.Value : _definition!.BasedOn;

        internal int? NumberId => _style != null
            ? _style.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value
            : _definition!.NumberId;

        internal int? Level => _style != null
            ? _style.StyleParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value
            : _definition!.Level;
    }

    /// <summary>
    /// Builds a memoizing catalog. Inside a read-only pass over the same document the pass shares one catalog,
    /// so styles and numbering are read once for the whole pass.
    /// </summary>
    internal static StyleCatalog CreateStyleCatalog(WordDocument document) {
        if (ActiveReadOnlyPass.Value is { } pass && ReferenceEquals(pass.Document, document)) {
            return pass.Catalog ??= new StyleCatalog(document._wordprocessingDocument.MainDocumentPart, memoizeResolutions: true);
        }
        return new StyleCatalog(document._wordprocessingDocument.MainDocumentPart, memoizeResolutions: true);
    }

    /// <summary>
    /// Shares one style catalog across a read-only pass over <paramref name="document"/>, such as building an
    /// inspection snapshot, so each paragraph's list resolution does not revalidate the cached catalog.
    /// The pass must not change styles or numbering while the scope is active.
    /// </summary>
    internal static IDisposable BeginReadOnlyPass(WordDocument document) {
        var scope = new ReadOnlyPassScope(document, ActiveReadOnlyPass.Value);
        ActiveReadOnlyPass.Value = scope;
        return scope;
    }

    private static readonly AsyncLocal<ReadOnlyPassScope?> ActiveReadOnlyPass = new();

    private sealed class ReadOnlyPassScope : IDisposable {
        private readonly ReadOnlyPassScope? _previous;
        private bool _disposed;

        internal ReadOnlyPassScope(WordDocument document, ReadOnlyPassScope? previous) {
            Document = document;
            _previous = previous;
        }

        internal WordDocument Document { get; }
        internal StyleCatalog? Catalog { get; set; }

        public void Dispose() {
            if (_disposed) return;
            _disposed = true;
            ActiveReadOnlyPass.Value = _previous;
        }
    }

    internal static StyleCatalog GetCachedStyleCatalog(WordDocument document) {
        if (ActiveReadOnlyPass.Value is { } pass && ReferenceEquals(pass.Document, document)) {
            return pass.Catalog ??= new StyleCatalog(document._wordprocessingDocument.MainDocumentPart, memoizeResolutions: true);
        }
        return GetUnscopedStyleCatalog(document);
    }

    private static StyleCatalog GetUnscopedStyleCatalog(WordDocument document) {
        MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
        if (mainPart == null) return new StyleCatalog(null);
        if (document._openXmlDocumentAccessed) return new StyleCatalog(mainPart);

        CachedStyleCatalog cached = StyleCatalogs.GetValue(mainPart, _ => new CachedStyleCatalog());
        lock (cached) {
            // An unloaded styles part is read as a stream and is not loaded just to validate the cache.
            StylesPart? stylesPart = mainPart.StyleDefinitionsPart;
            StylesPart? effectsPart = mainPart.StylesWithEffectsPart;
            Styles? styles = LoadedStyles(stylesPart);
            Styles? effects = LoadedStyles(effectsPart);
            Numbering? numbering = mainPart.NumberingDefinitionsPart?.Numbering;
            int styleCount = styles?.ChildElements.Count ?? 0;
            int effectCount = effects?.ChildElements.Count ?? 0;
            int numberingCount = numbering?.ChildElements.Count ?? 0;
            int numberingRevision = numbering == null ? 0 : Volatile.Read(ref NumberingRevisions.GetValue(numbering, _ => new NumberingRevision()).Value);
            if (cached.Catalog == null ||
                !ReferenceEquals(cached.StylesPart, stylesPart) || !ReferenceEquals(cached.StylesWithEffectsPart, effectsPart) ||
                !ReferenceEquals(cached.Styles, styles) || cached.StyleCount != styleCount ||
                !ReferenceEquals(cached.StylesWithEffects, effects) || cached.EffectStyleCount != effectCount ||
                !ReferenceEquals(cached.Numbering, numbering) || cached.NumberingCount != numberingCount || cached.NumberingRevision != numberingRevision) {
                cached.Catalog = new StyleCatalog(mainPart);
                // Building the catalog loads a styles part that could not be streamed; record what it read.
                styles = LoadedStyles(stylesPart);
                effects = LoadedStyles(effectsPart);
                styleCount = styles?.ChildElements.Count ?? 0;
                effectCount = effects?.ChildElements.Count ?? 0;
                cached.StylesPart = stylesPart;
                cached.StylesWithEffectsPart = effectsPart;
                cached.Styles = styles;
                cached.StylesWithEffects = effects;
                cached.Numbering = numbering;
                cached.StyleCount = styleCount;
                cached.EffectStyleCount = effectCount;
                cached.NumberingCount = numberingCount;
                cached.NumberingRevision = numberingRevision;
            }
            return cached.Catalog;
        }
    }

    private static Styles? LoadedStyles(StylesPart? part) =>
        part != null && part.IsRootElementLoaded ? part.Styles : null;

    internal static void InvalidateNumbering(Numbering? numbering) {
        if (numbering == null) return;
        Interlocked.Increment(ref NumberingRevisions.GetValue(numbering, _ => new NumberingRevision()).Value);
    }

    internal readonly struct ResolvedNumbering {
        internal ResolvedNumbering(int numberId, int level) {
            NumberId = numberId;
            Level = level;
        }

        internal int NumberId { get; }
        internal int Level { get; }
    }

    internal static bool TryResolve(WordParagraph paragraph, out ResolvedNumbering numbering, StyleCatalog? styleCatalog = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        numbering = default;
        if (paragraph?._paragraph == null || paragraph._document == null) {
            return false;
        }

        NumberingProperties? direct = paragraph._paragraph.ParagraphProperties?.NumberingProperties;
        int? numberId = ReadNumberId(direct);
        int? level = ReadLevel(direct);
        bool hasDirectNumberId = numberId.HasValue;
        bool hasDirectLevel = level.HasValue;
        if (numberId == 0) {
            return false;
        }

        // The common direct-numbering path needs no style catalog lookup.
        if (numberId > 0 && level.HasValue) {
            numbering = new ResolvedNumbering(numberId.Value, Math.Max(0, level.Value));
            return true;
        }

        if (!numberId.HasValue || !level.HasValue) {
            styleCatalog ??= GetCachedStyleCatalog(paragraph._document);
            string? styleId = paragraph._paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            NumberingProperties? inherited = ResolveStyleNumbering(styleCatalog, styleId, cancellationToken);
            numberId ??= ReadNumberId(inherited);
            if (!numberId.HasValue && ResolveLinkedStyle(styleCatalog, styleId, cancellationToken) is { } linked) {
                numberId = linked.NumberId;
                level ??= linked.Level;
            }
            if (numberId > 0) {
                int? linkedLevel = ResolveLinkedLevel(styleCatalog, numberId.Value, styleId, cancellationToken);
                if (!hasDirectNumberId && !hasDirectLevel && linkedLevel.HasValue) {
                    level = linkedLevel.Value;
                } else if (!level.HasValue) {
                    level = linkedLevel ?? ReadLevel(inherited);
                }
            }
        }

        if (!numberId.HasValue || numberId.Value <= 0) {
            return false;
        }

        numbering = new ResolvedNumbering(numberId.Value, Math.Max(0, level ?? 0));
        return true;
    }

    private static NumberingProperties? ResolveStyleNumbering(StyleCatalog catalog, string? styleId, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(styleId)) styleId = catalog.DefaultStyleId;
        if (string.IsNullOrWhiteSpace(styleId)) return null;
        if (catalog.NumberingByStyle?.TryGetValue(styleId!, out NumberingProperties? cached) == true) return cached;

        int? numberId = null;
        int? level = null;
        string? currentStyleId = styleId;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visited.Count > MaximumStyleInheritanceDepth) throw new InvalidDataException("Word style inheritance exceeds the supported depth.");
            if (!catalog.ById.TryGetValue(currentStyleId!, out ParagraphStyleEntry? style)) {
                break;
            }

            numberId ??= style.NumberId;
            level ??= style.Level;
            if (numberId.HasValue && level.HasValue) {
                break;
            }

            currentStyleId = style.BasedOn;
        }

        if (!numberId.HasValue && !level.HasValue) {
            if (catalog.NumberingByStyle != null) catalog.NumberingByStyle[styleId!] = null;
            return null;
        }

        var resolved = new NumberingProperties();
        if (level.HasValue) {
            resolved.Append(new NumberingLevelReference { Val = level.Value });
        }
        if (numberId.HasValue) {
            resolved.Append(new NumberingId { Val = numberId.Value });
        }
        if (catalog.NumberingByStyle != null) catalog.NumberingByStyle[styleId!] = resolved;
        return resolved;
    }

    private static int? ResolveLinkedLevel(StyleCatalog catalog, int numberId, string? styleId, CancellationToken cancellationToken) {
        string? currentStyleId = string.IsNullOrWhiteSpace(styleId) ? catalog.DefaultStyleId : styleId;
        if (string.IsNullOrWhiteSpace(currentStyleId)) return null;
        (int NumberId, string StyleId) key = (numberId, currentStyleId!);
        if (catalog.LinkedLevelByStyle?.TryGetValue(key, out int? cached) == true) return cached;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visited.Count > MaximumStyleInheritanceDepth) throw new InvalidDataException("Word style inheritance exceeds the supported depth.");
            if (catalog.LinkedLevels.TryGetValue((numberId, currentStyleId!), out int level)) {
                if (catalog.LinkedLevelByStyle != null) catalog.LinkedLevelByStyle[key] = level;
                return level;
            }
            currentStyleId = catalog.ById.TryGetValue(currentStyleId!, out ParagraphStyleEntry? style)
                ? style.BasedOn
                : null;
        }
        if (catalog.LinkedLevelByStyle != null) catalog.LinkedLevelByStyle[key] = null;
        return null;
    }

    private static ResolvedNumbering? ResolveLinkedStyle(StyleCatalog catalog, string? styleId, CancellationToken cancellationToken) {
        string? currentStyleId = string.IsNullOrWhiteSpace(styleId) ? catalog.DefaultStyleId : styleId;
        if (string.IsNullOrWhiteSpace(currentStyleId)) return null;
        string cacheKey = currentStyleId!;
        if (catalog.LinkedStyleByStyle?.TryGetValue(cacheKey, out ResolvedNumbering? cached) == true) return cached;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visited.Count > MaximumStyleInheritanceDepth) throw new InvalidDataException("Word style inheritance exceeds the supported depth.");
            if (catalog.LinkedStyles.TryGetValue(currentStyleId!, out ResolvedNumbering numbering)) {
                if (catalog.LinkedStyleByStyle != null) catalog.LinkedStyleByStyle[cacheKey] = numbering;
                return numbering;
            }
            currentStyleId = catalog.ById.TryGetValue(currentStyleId!, out ParagraphStyleEntry? style)
                ? style.BasedOn
                : null;
        }
        if (catalog.LinkedStyleByStyle != null) catalog.LinkedStyleByStyle[cacheKey] = null;
        return null;
    }

    private static int? ReadNumberId(NumberingProperties? properties) =>
        properties?.NumberingId?.Val?.Value;

    private static int? ReadLevel(NumberingProperties? properties) =>
        properties?.NumberingLevelReference?.Val?.Value;
}

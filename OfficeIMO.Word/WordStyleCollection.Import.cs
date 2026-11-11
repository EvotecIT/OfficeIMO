using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Decides what happens when an imported style matches a style the target document already has.</summary>
/// <remarks>A style matches when it has the same type and name, or the same type and identifier.</remarks>
public enum WordStyleConflictMode {
    /// <summary>Keep the target's style. Imported styles that refer to it use the target's definition.</summary>
    KeepTarget,
    /// <summary>Replace the target's definition with the imported one, keeping the target's identifier.</summary>
    Overwrite,
    /// <summary>Import the style alongside the target's under a new, unique identifier and name.</summary>
    Rename
}

/// <summary>What <see cref="WordStyleCollection.Import(WordDocument, WordStyleConflictMode, IEnumerable{string}?)"/> did with each style.</summary>
public sealed class WordStyleImportReport {
    internal WordStyleImportReport(IReadOnlyList<string> added, IReadOnlyList<string> replaced, IReadOnlyList<string> kept,
        IReadOnlyDictionary<string, string> renamed) {
        Added = added;
        Replaced = replaced;
        Kept = kept;
        Renamed = renamed;
    }

    /// <summary>Gets the identifiers of styles added under their source identifier.</summary>
    public IReadOnlyList<string> Added { get; }

    /// <summary>Gets the identifiers of target styles whose definitions were replaced.</summary>
    public IReadOnlyList<string> Replaced { get; }

    /// <summary>Gets the identifiers of target styles kept in place of a matching imported style.</summary>
    public IReadOnlyList<string> Kept { get; }

    /// <summary>Gets each source identifier that was imported under a new identifier, mapped to that identifier.</summary>
    public IReadOnlyDictionary<string, string> Renamed { get; }
}

public sealed partial class WordStyleCollection {
    /// <summary>
    /// Copies styles from another document, together with the styles they are based on, link to or are followed by,
    /// and the lists they number paragraphs with.
    /// </summary>
    /// <param name="source">The document or template to copy from. It is not changed; a read-only load is enough.</param>
    /// <param name="conflictMode">
    /// What to do when the target already has a style matching a requested style. Styles copied only because a
    /// requested style depends on them use the target's matching style when there is one.
    /// </param>
    /// <param name="styleIds">Identifiers of the styles to copy; <see langword="null"/> copies every style.</param>
    /// <returns>Which styles were added, replaced, kept or renamed.</returns>
    /// <remarks>
    /// A target style that matches only by identifier but has a different type is never overwritten or kept; the
    /// imported style is renamed. Document defaults and the document theme are not copied.
    /// </remarks>
    /// <exception cref="KeyNotFoundException">A requested identifier does not exist in <paramref name="source"/>.</exception>
    public WordStyleImportReport Import(WordDocument source, WordStyleConflictMode conflictMode = WordStyleConflictMode.KeepTarget,
        IEnumerable<string>? styleIds = null) {
        ArgumentNullException.ThrowIfNull(source);
        if (ReferenceEquals(source, _document)) throw new ArgumentException("A document cannot import its own styles.", nameof(source));
        Styles targetStyles = EnsureStyles();
        Styles? sourceStyles = source._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        var plan = new ImportPlan(sourceStyles, targetStyles, conflictMode);
        if (sourceStyles != null) {
            plan.Select(styleIds);
            plan.Commit(source._wordprocessingDocument!.MainDocumentPart!, _document._wordprocessingDocument.MainDocumentPart!);
            // Lists created after the import must not reuse the imported list definitions' identifiers.
            WordListStyles.InitializeAbstractNumberId(_document._wordprocessingDocument);
        } else if (styleIds?.Any() == true) {
            throw new KeyNotFoundException("The source document defines no styles.");
        }
        return plan.CreateReport();
    }

    private sealed class ImportPlan {
        private readonly Styles? _source;
        private readonly Styles _target;
        private readonly WordStyleConflictMode _mode;
        private readonly Dictionary<string, Style> _sourceById = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _map = new(StringComparer.Ordinal);
        private readonly List<(Style Source, string Id, string? Name, Style? Replaces)> _copies = new();
        private readonly List<string> _added = new(), _replaced = new(), _kept = new();
        private readonly Dictionary<string, string> _renamed = new(StringComparer.Ordinal);
        private readonly Dictionary<int, int> _numberingMap = new();

        internal ImportPlan(Styles? source, Styles target, WordStyleConflictMode mode) {
            _source = source;
            _target = target;
            _mode = mode;
            if (source == null) return;
            foreach (Style style in source.Elements<Style>()) {
                if (style.StyleId?.Value is { Length: > 0 } id && !_sourceById.ContainsKey(id)) _sourceById.Add(id, style);
            }
        }

        internal void Select(IEnumerable<string>? requested) {
            var selected = new HashSet<string>(StringComparer.Ordinal);
            var explicitlyRequested = new HashSet<string>(StringComparer.Ordinal);
            var pending = new Queue<string>();
            void Include(string id) { if (selected.Add(id)) pending.Enqueue(id); }
            if (requested == null) {
                foreach (string id in _sourceById.Keys) Include(id);
            } else {
                foreach (string id in requested) {
                    if (!_sourceById.ContainsKey(id)) throw new KeyNotFoundException($"The source document defines no style with the identifier '{id}'.");
                    explicitlyRequested.Add(id);
                    Include(id);
                }
            }
            while (pending.Count > 0) {
                Style style = _sourceById[pending.Dequeue()];
                foreach (string? related in new[] { style.BasedOn?.Val?.Value, style.LinkedStyle?.Val?.Value, style.NextParagraphStyle?.Val?.Value }) {
                    if (!string.IsNullOrEmpty(related) && _sourceById.ContainsKey(related!)) Include(related!);
                }
            }

            var targetIds = new HashSet<string>(_target.Elements<Style>().Select(style => style.StyleId?.Value).OfType<string>(), StringComparer.Ordinal);
            // Selected source styles keep their identifiers when added, so a renamed copy must not take one of them.
            targetIds.UnionWith(selected);
            var targetNames = new HashSet<string>(_target.Elements<Style>().Select(style => style.StyleName?.Val?.Value?.Trim()).OfType<string>(), StringComparer.OrdinalIgnoreCase);
            foreach (Style style in _source!.Elements<Style>()) {
                string id = style.StyleId?.Value ?? string.Empty;
                if (!selected.Contains(id) || _map.ContainsKey(id)) continue;
                Style? match = FindMatch(style, out bool typeClash);
                // Styles pulled in only as dependencies never replace or duplicate a matching target style.
                WordStyleConflictMode mode = requested == null || explicitlyRequested.Contains(id) ? _mode : WordStyleConflictMode.KeepTarget;
                // A target style of another type with the same name is a clash too: names are unique across types.
                string? sourceName = style.StyleName?.Val?.Value?.Trim();
                if (match == null && sourceName != null && targetNames.Contains(sourceName)) typeClash = true;
                if (match == null && !typeClash) {
                    _map[id] = id;
                    targetIds.Add(id);
                    if (style.StyleName?.Val?.Value is { } name) targetNames.Add(name.Trim());
                    _copies.Add((style, id, null, null));
                    _added.Add(id);
                } else if (match != null && mode == WordStyleConflictMode.KeepTarget) {
                    _map[id] = match.StyleId!.Value!;
                    _kept.Add(match.StyleId!.Value!);
                } else if (match != null && mode == WordStyleConflictMode.Overwrite) {
                    _map[id] = match.StyleId!.Value!;
                    _copies.Add((style, match.StyleId!.Value!, null, match));
                    _replaced.Add(match.StyleId!.Value!);
                } else {
                    (string newId, string newName) = CreateUniqueIdentity(style, targetIds, targetNames);
                    _map[id] = newId;
                    _copies.Add((style, newId, newName, null));
                    _renamed.Add(id, newId);
                }
            }
        }

        private Style? FindMatch(Style source, out bool typeClash) {
            typeClash = false;
            StyleValues? type = source.Type?.Value;
            string? name = source.StyleName?.Val?.Value?.Trim();
            Style? byName = name == null ? null : _target.Elements<Style>().FirstOrDefault(style =>
                style.Type?.Value == type && string.Equals(style.StyleName?.Val?.Value?.Trim(), name, StringComparison.OrdinalIgnoreCase));
            if (byName != null) return byName;
            Style? byId = _target.Elements<Style>().FirstOrDefault(style => string.Equals(style.StyleId?.Value, source.StyleId?.Value, StringComparison.Ordinal));
            if (byId == null) return null;
            if (byId.Type?.Value == type) return byId;
            typeClash = true;
            return null;
        }

        private static (string Id, string Name) CreateUniqueIdentity(Style style, HashSet<string> ids, HashSet<string> names) {
            string baseId = style.StyleId!.Value!;
            string baseName = style.StyleName?.Val?.Value?.Trim() ?? baseId;
            for (int suffix = 1; ; suffix++) {
                string id = baseId + "_" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
                string name = baseName + " " + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (!ids.Contains(id) && !names.Contains(name)) {
                    ids.Add(id);
                    names.Add(name);
                    return (id, name);
                }
            }
        }

        internal void Commit(MainDocumentPart sourceMain, MainDocumentPart targetMain) {
            ImportNumbering(sourceMain, targetMain);
            foreach ((Style source, string id, string? name, Style? replaces) in _copies) {
                var clone = (Style)source.CloneNode(true);
                clone.StyleId = id;
                if (name != null) {
                    clone.StyleName = new StyleName { Val = name };
                    clone.CustomStyle = true;
                }
                if (clone.BasedOn?.Val?.Value is { } basedOn) clone.BasedOn.Val = Map(basedOn);
                if (clone.NextParagraphStyle?.Val?.Value is { } next) clone.NextParagraphStyle.Val = Map(next);
                if (clone.LinkedStyle?.Val?.Value is { } linked) clone.LinkedStyle.Val = Map(linked);
                foreach (NumberingId numbering in clone.Descendants<NumberingId>()) {
                    if (numbering.Val?.Value is int numberId && _numberingMap.TryGetValue(numberId, out int mapped)) numbering.Val = mapped;
                }
                if (replaces != null) {
                    // The target decides which style is its default; imported definitions do not move that choice.
                    clone.Default = replaces.Default?.Value;
                    replaces.InsertAfterSelf(clone);
                    replaces.Remove();
                } else {
                    clone.Default = null;
                    _target.Append(clone);
                }
            }
            WordDocumentMerge.MergeNamespaces(sourceMain.StyleDefinitionsPart?.Styles, targetMain.StyleDefinitionsPart?.Styles);
        }

        private string Map(string sourceId) => _map.TryGetValue(sourceId, out string? mapped) ? mapped : sourceId;

        /// <summary>Copies only the list definitions that copied styles number with.</summary>
        private void ImportNumbering(MainDocumentPart sourceMain, MainDocumentPart targetMain) {
            Numbering? sourceNumbering = sourceMain.NumberingDefinitionsPart?.Numbering;
            if (sourceNumbering == null) return;
            var used = new HashSet<int>(_copies.SelectMany(copy => copy.Source.Descendants<NumberingId>())
                .Select(numbering => numbering.Val?.Value ?? 0).Where(id => id > 0));
            if (used.Count == 0) return;

            NumberingDefinitionsPart targetPart = targetMain.NumberingDefinitionsPart ?? targetMain.AddNewPart<NumberingDefinitionsPart>();
            Numbering target = targetPart.Numbering ??= new Numbering();
            int nextAbstract = target.Elements<AbstractNum>().Select(item => item.AbstractNumberId?.Value ?? -1).DefaultIfEmpty(-1).Max() + 1;
            int nextNumber = target.Elements<NumberingInstance>().Select(item => item.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
            var abstractMap = new Dictionary<int, int>();
            OpenXmlElement? abstractAnchor = target.Elements<NumberingInstance>().FirstOrDefault()
                ?? (OpenXmlElement?)target.Elements<NumberingIdMacAtCleanup>().FirstOrDefault();

            foreach (NumberingInstance instance in sourceNumbering.Elements<NumberingInstance>()) {
                if (instance.NumberID?.Value is not int numberId || !used.Contains(numberId)) continue;
                int? sourceAbstract = instance.AbstractNumId?.Val?.Value;
                if (sourceAbstract is int abstractId && !abstractMap.ContainsKey(abstractId)) {
                    AbstractNum? definition = sourceNumbering.Elements<AbstractNum>().FirstOrDefault(item => item.AbstractNumberId?.Value == abstractId);
                    if (definition != null) {
                        var abstractClone = (AbstractNum)definition.CloneNode(true);
                        abstractClone.AbstractNumberId = nextAbstract;
                        abstractMap[abstractId] = nextAbstract++;
                        // Picture bullets live in the source package; fall back to the text bullet.
                        foreach (LevelPictureBulletId picture in abstractClone.Descendants<LevelPictureBulletId>().ToList()) picture.Remove();
                        foreach (ParagraphStyleId levelStyle in abstractClone.Descendants<ParagraphStyleId>()) {
                            if (levelStyle.Val?.Value is { } styleId) levelStyle.Val = Map(styleId);
                        }
                        if (abstractAnchor != null) target.InsertBefore(abstractClone, abstractAnchor);
                        else target.Append(abstractClone);
                    }
                }
                var instanceClone = (NumberingInstance)instance.CloneNode(true);
                instanceClone.NumberID = nextNumber;
                _numberingMap[numberId] = nextNumber++;
                if (instanceClone.AbstractNumId?.Val?.Value is int oldAbstract && abstractMap.TryGetValue(oldAbstract, out int newAbstract))
                    instanceClone.AbstractNumId.Val = newAbstract;
                OpenXmlElement? numberAnchor = target.Elements<NumberingIdMacAtCleanup>().FirstOrDefault();
                if (numberAnchor != null) target.InsertBefore(instanceClone, numberAnchor);
                else target.Append(instanceClone);
            }
            WordDocumentMerge.MergeNamespaces(sourceNumbering, target);
        }

        internal WordStyleImportReport CreateReport() => new(_added, _replaced, _kept, _renamed);
    }
}

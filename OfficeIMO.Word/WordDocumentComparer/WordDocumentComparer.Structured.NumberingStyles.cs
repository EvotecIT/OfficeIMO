using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static OfficeXmlNode ResolveParagraphNumberingProperties(
            OfficeXmlNode paragraph,
            ParagraphNumberingStyleCatalog styleCatalog) {
            _ = TryResolveParagraphNumberingProperties(
                paragraph,
                styleCatalog,
                comparisonWorkBudget: null,
                out OfficeXmlNode numbering);
            return numbering;
        }

        private static bool TryResolveParagraphNumberingProperties(
            OfficeXmlNode paragraph,
            ParagraphNumberingStyleCatalog styleCatalog,
            ComparisonWorkBudget? comparisonWorkBudget,
            out OfficeXmlNode numbering) {
            OfficeXmlNode directNumbering = paragraph.Child(W, "pPr").Child(W, "numPr");
            if (!directNumbering.IsNone) {
                numbering = directNumbering;
                return true;
            }
            string? styleId = ChildVal(paragraph.Child(W, "pPr"), "pStyle");
            return styleCatalog.TryResolveNumbering(styleId, comparisonWorkBudget, out numbering);
        }

        /// <summary>Caches the combined style definitions used while resolving paragraph numbering for one analysis.</summary>
        private sealed class ParagraphNumberingStyleCatalog {
            private const string DefaultParagraphStylesCacheKey = "\0default-paragraph-styles";
            private readonly Dictionary<string, OfficeXmlNode> _resolvedNumbering =
                new(StringComparer.Ordinal);

            private ParagraphNumberingStyleCatalog(
                ILookup<string, ComparisonStyle> stylesById,
                IReadOnlyList<string> defaultParagraphStyleIds) {
                StylesById = stylesById;
                DefaultParagraphStyleIds = defaultParagraphStyleIds;
            }

            internal ILookup<string, ComparisonStyle> StylesById { get; }
            internal IReadOnlyList<string> DefaultParagraphStyleIds { get; }

            internal OfficeXmlNode ResolveNumbering(string? styleId) {
                _ = TryResolveNumbering(styleId, comparisonWorkBudget: null, out OfficeXmlNode numbering);
                return numbering;
            }

            internal bool TryResolveNumbering(
                string? styleId,
                ComparisonWorkBudget? comparisonWorkBudget,
                out OfficeXmlNode numbering) {
                string cacheKey = string.IsNullOrWhiteSpace(styleId)
                    ? DefaultParagraphStylesCacheKey
                    : styleId!;
                if (_resolvedNumbering.TryGetValue(cacheKey, out OfficeXmlNode cached)) {
                    numbering = cached;
                    return true;
                }

                var visiting = new HashSet<string>(StringComparer.Ordinal);
                if (cacheKey == DefaultParagraphStylesCacheKey) {
                    foreach (string defaultStyleId in DefaultParagraphStyleIds) {
                        if (!TryResolveStyleNumbering(
                                defaultStyleId,
                                comparisonWorkBudget,
                                visiting,
                                out OfficeXmlNode resolved)) {
                            numbering = default;
                            return false;
                        }
                        if (!resolved.IsNone) {
                            _resolvedNumbering[cacheKey] = resolved;
                            numbering = resolved;
                            return true;
                        }
                    }
                    _resolvedNumbering[cacheKey] = default;
                    numbering = default;
                    return true;
                }

                return TryResolveStyleNumbering(cacheKey, comparisonWorkBudget, visiting, out numbering);
            }

            private bool TryResolveStyleNumbering(
                string styleId,
                ComparisonWorkBudget? comparisonWorkBudget,
                HashSet<string> visiting,
                out OfficeXmlNode numbering) {
                if (_resolvedNumbering.TryGetValue(styleId, out OfficeXmlNode cached)) {
                    numbering = cached;
                    return true;
                }
                var frames = new List<StyleNumberingResolutionFrame>();
                if (!TryPushStyleResolutionFrame(
                        styleId,
                        comparisonWorkBudget,
                        visiting,
                        frames)) {
                    numbering = default;
                    return false;
                }

                while (frames.Count > 0) {
                    StyleNumberingResolutionFrame frame = frames[frames.Count - 1];
                    if (frame.NextStyleIndex >= frame.Styles.Count) {
                        CompleteStyleResolution(frames, visiting, frame.StyleId, default);
                        continue;
                    }

                    ComparisonStyle style = frame.Styles[frame.NextStyleIndex++];
                    OfficeXmlNode direct = style.Element.Child(W, "pPr").Child(W, "numPr");
                    if (!direct.IsNone) {
                        CompleteStyleResolution(frames, visiting, frame.StyleId, direct);
                        continue;
                    }

                    string? baseStyleId = style.BasedOn;
                    if (string.IsNullOrWhiteSpace(baseStyleId) || visiting.Contains(baseStyleId!)) continue;
                    if (_resolvedNumbering.TryGetValue(baseStyleId!, out OfficeXmlNode inherited)) {
                        if (!inherited.IsNone) {
                            CompleteStyleResolution(frames, visiting, frame.StyleId, inherited);
                        }
                        continue;
                    }
                    if (!TryPushStyleResolutionFrame(
                            baseStyleId!,
                            comparisonWorkBudget,
                            visiting,
                            frames)) {
                        foreach (StyleNumberingResolutionFrame pending in frames) visiting.Remove(pending.StyleId);
                        numbering = default;
                        return false;
                    }
                }

                numbering = _resolvedNumbering[styleId];
                return true;
            }

            private bool TryPushStyleResolutionFrame(
                string styleId,
                ComparisonWorkBudget? comparisonWorkBudget,
                ISet<string> visiting,
                ICollection<StyleNumberingResolutionFrame> frames) {
                if (!visiting.Add(styleId)) return true;
                if (comparisonWorkBudget != null && !comparisonWorkBudget.TryConsume(1)) {
                    visiting.Remove(styleId);
                    return false;
                }
                frames.Add(new StyleNumberingResolutionFrame(styleId, StylesById[styleId].ToArray()));
                return true;
            }

            private void CompleteStyleResolution(
                IList<StyleNumberingResolutionFrame> frames,
                ISet<string> visiting,
                string styleId,
                OfficeXmlNode numbering) {
                _resolvedNumbering[styleId] = numbering;
                visiting.Remove(styleId);
                frames.RemoveAt(frames.Count - 1);
                while (!numbering.IsNone && frames.Count > 0) {
                    StyleNumberingResolutionFrame parent = frames[frames.Count - 1];
                    _resolvedNumbering[parent.StyleId] = numbering;
                    visiting.Remove(parent.StyleId);
                    frames.RemoveAt(frames.Count - 1);
                }
            }

            private sealed class StyleNumberingResolutionFrame {
                internal StyleNumberingResolutionFrame(string styleId, IReadOnlyList<ComparisonStyle> styles) {
                    StyleId = styleId;
                    Styles = styles;
                }

                internal string StyleId { get; }
                internal IReadOnlyList<ComparisonStyle> Styles { get; }
                internal int NextStyleIndex { get; set; }
            }

            internal static ParagraphNumberingStyleCatalog? Create(
                WordPartView mainPart,
                ComparisonWorkBudget? comparisonWorkBudget) {
                var styles = new List<ComparisonStyle>();
                foreach (ComparisonStyle style in EnumerateComparisonStyles(mainPart)) {
                    if (comparisonWorkBudget != null && !comparisonWorkBudget.TryConsume(1)) return null;
                    styles.Add(style);
                }
                ILookup<string, ComparisonStyle> stylesById = styles
                    .Where(style => !string.IsNullOrWhiteSpace(style.StyleId))
                    .ToLookup(style => style.StyleId!, StringComparer.Ordinal);
                string[] defaultParagraphStyleIds = styles
                    .Where(style =>
                        style.IsParagraphStyle &&
                        style.IsDefault &&
                        !string.IsNullOrWhiteSpace(style.StyleId))
                    .Select(style => style.StyleId!)
                    .ToArray();
                return new ParagraphNumberingStyleCatalog(stylesById, defaultParagraphStyleIds);
            }
        }

        /// <summary>Reuses numbering style catalogs across all feature and limitation passes in one comparison.</summary>
        private sealed class ParagraphNumberingStyleCatalogCache {
            private readonly Dictionary<WordPartView, ParagraphNumberingStyleCatalog> _catalogs = new();

            internal ParagraphNumberingStyleCatalog GetOrCreate(WordPartView mainPart) {
                if (_catalogs.TryGetValue(mainPart, out ParagraphNumberingStyleCatalog? catalog)) return catalog;
                catalog = ParagraphNumberingStyleCatalog.Create(mainPart, comparisonWorkBudget: null)!;
                _catalogs.Add(mainPart, catalog);
                return catalog;
            }

            internal bool TryGetOrCreate(
                WordPartView mainPart,
                ComparisonWorkBudget comparisonWorkBudget,
                out ParagraphNumberingStyleCatalog? catalog) {
                if (_catalogs.TryGetValue(mainPart, out catalog)) return true;
                catalog = ParagraphNumberingStyleCatalog.Create(mainPart, comparisonWorkBudget);
                if (catalog == null) return false;
                _catalogs.Add(mainPart, catalog);
                return true;
            }
        }
    }
}

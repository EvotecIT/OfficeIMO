using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static NumberingProperties? ResolveParagraphNumberingProperties(
            Paragraph paragraph,
            ParagraphNumberingStyleCatalog styleCatalog) {
            _ = TryResolveParagraphNumberingProperties(
                paragraph,
                styleCatalog,
                comparisonWorkBudget: null,
                out NumberingProperties? numbering);
            return numbering;
        }

        private static bool TryResolveParagraphNumberingProperties(
            Paragraph paragraph,
            ParagraphNumberingStyleCatalog styleCatalog,
            ComparisonWorkBudget? comparisonWorkBudget,
            out NumberingProperties? numbering) {
            NumberingProperties? directNumbering = paragraph.ParagraphProperties?.NumberingProperties;
            if (directNumbering != null) {
                numbering = directNumbering;
                return true;
            }
            string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            return styleCatalog.TryResolveNumbering(styleId, comparisonWorkBudget, out numbering);
        }

        /// <summary>Caches the combined style definitions used while resolving paragraph numbering for one analysis.</summary>
        private sealed class ParagraphNumberingStyleCatalog {
            private const string DefaultParagraphStylesCacheKey = "\0default-paragraph-styles";
            private readonly Dictionary<string, NumberingProperties?> _resolvedNumbering =
                new(StringComparer.Ordinal);

            private ParagraphNumberingStyleCatalog(
                ILookup<string, Style> stylesById,
                IReadOnlyList<string> defaultParagraphStyleIds) {
                StylesById = stylesById;
                DefaultParagraphStyleIds = defaultParagraphStyleIds;
            }

            internal ILookup<string, Style> StylesById { get; }
            internal IReadOnlyList<string> DefaultParagraphStyleIds { get; }

            internal NumberingProperties? ResolveNumbering(string? styleId) {
                _ = TryResolveNumbering(styleId, comparisonWorkBudget: null, out NumberingProperties? numbering);
                return numbering;
            }

            internal bool TryResolveNumbering(
                string? styleId,
                ComparisonWorkBudget? comparisonWorkBudget,
                out NumberingProperties? numbering) {
                string cacheKey = string.IsNullOrWhiteSpace(styleId)
                    ? DefaultParagraphStylesCacheKey
                    : styleId!;
                if (_resolvedNumbering.TryGetValue(cacheKey, out NumberingProperties? cached)) {
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
                                out NumberingProperties? resolved)) {
                            numbering = null;
                            return false;
                        }
                        if (resolved != null) {
                            _resolvedNumbering[cacheKey] = resolved;
                            numbering = resolved;
                            return true;
                        }
                    }
                    _resolvedNumbering[cacheKey] = null;
                    numbering = null;
                    return true;
                }

                return TryResolveStyleNumbering(cacheKey, comparisonWorkBudget, visiting, out numbering);
            }

            private bool TryResolveStyleNumbering(
                string styleId,
                ComparisonWorkBudget? comparisonWorkBudget,
                HashSet<string> visiting,
                out NumberingProperties? numbering) {
                if (_resolvedNumbering.TryGetValue(styleId, out NumberingProperties? cached)) {
                    numbering = cached;
                    return true;
                }
                var frames = new List<StyleNumberingResolutionFrame>();
                if (!TryPushStyleResolutionFrame(
                        styleId,
                        comparisonWorkBudget,
                        visiting,
                        frames)) {
                    numbering = null;
                    return false;
                }

                while (frames.Count > 0) {
                    StyleNumberingResolutionFrame frame = frames[frames.Count - 1];
                    if (frame.NextStyleIndex >= frame.Styles.Count) {
                        CompleteStyleResolution(frames, visiting, frame.StyleId, null);
                        continue;
                    }

                    Style style = frame.Styles[frame.NextStyleIndex++];
                    NumberingProperties? direct = style.StyleParagraphProperties?.NumberingProperties;
                    if (direct != null) {
                        CompleteStyleResolution(frames, visiting, frame.StyleId, direct);
                        continue;
                    }

                    string? baseStyleId = style.BasedOn?.Val?.Value;
                    if (string.IsNullOrWhiteSpace(baseStyleId) || visiting.Contains(baseStyleId!)) continue;
                    if (_resolvedNumbering.TryGetValue(baseStyleId!, out NumberingProperties? inherited)) {
                        if (inherited != null) {
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
                        numbering = null;
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
                NumberingProperties? numbering) {
                _resolvedNumbering[styleId] = numbering;
                visiting.Remove(styleId);
                frames.RemoveAt(frames.Count - 1);
                while (numbering != null && frames.Count > 0) {
                    StyleNumberingResolutionFrame parent = frames[frames.Count - 1];
                    _resolvedNumbering[parent.StyleId] = numbering;
                    visiting.Remove(parent.StyleId);
                    frames.RemoveAt(frames.Count - 1);
                }
            }

            private sealed class StyleNumberingResolutionFrame {
                internal StyleNumberingResolutionFrame(string styleId, IReadOnlyList<Style> styles) {
                    StyleId = styleId;
                    Styles = styles;
                }

                internal string StyleId { get; }
                internal IReadOnlyList<Style> Styles { get; }
                internal int NextStyleIndex { get; set; }
            }

            internal static ParagraphNumberingStyleCatalog? Create(
                MainDocumentPart mainPart,
                ComparisonWorkBudget? comparisonWorkBudget) {
                var styles = new List<Style>();
                IEnumerable<Style> candidates =
                    (mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                    .Concat(mainPart.StylesWithEffectsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>());
                foreach (Style style in candidates) {
                    if (comparisonWorkBudget != null && !comparisonWorkBudget.TryConsume(1)) return null;
                    styles.Add(style);
                }
                ILookup<string, Style> stylesById = styles
                    .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                    .ToLookup(style => style.StyleId!.Value!, StringComparer.Ordinal);
                string[] defaultParagraphStyleIds = styles
                    .Where(style =>
                        style.Type?.Value == StyleValues.Paragraph &&
                        style.Default?.Value == true &&
                        !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                    .Select(style => style.StyleId!.Value!)
                    .ToArray();
                return new ParagraphNumberingStyleCatalog(stylesById, defaultParagraphStyleIds);
            }
        }

        /// <summary>Reuses numbering style catalogs across all feature and limitation passes in one comparison.</summary>
        private sealed class ParagraphNumberingStyleCatalogCache {
            private readonly Dictionary<MainDocumentPart, ParagraphNumberingStyleCatalog> _catalogs = new();

            internal ParagraphNumberingStyleCatalog GetOrCreate(MainDocumentPart mainPart) {
                if (_catalogs.TryGetValue(mainPart, out ParagraphNumberingStyleCatalog? catalog)) return catalog;
                catalog = ParagraphNumberingStyleCatalog.Create(mainPart, comparisonWorkBudget: null)!;
                _catalogs.Add(mainPart, catalog);
                return catalog;
            }

            internal bool TryGetOrCreate(
                MainDocumentPart mainPart,
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

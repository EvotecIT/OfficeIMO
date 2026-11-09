using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private const uint FirstCustomNumberFormatId = 164U;

        // The attributes a cell format key reads through the target's indexes rather than as text.
        private static readonly string[] CellFormatReferenceAttributes = { "numFmtId", "fontId", "fillId", "borderId", "xfId" };

        /// <summary>
        /// Merges a source workbook's styles into the target workbook for one worksheet copy. Each source entry maps to the
        /// first target entry with the same properties and color binding, and is appended only when the target has none, so copying more sheets
        /// from the same source, or from a source with the same styles, adds nothing. A theme or palette reference has the
        /// same binding as the same reference in the target when both workbooks resolve it to the same color. The target is
        /// indexed afresh for every copy, so changes to either stylesheet between copies are always seen, and entries are
        /// visited in index order, so equal inputs give equal output.
        /// </summary>
        private sealed class StylesheetMerger {
            private readonly Stylesheet _source;
            private readonly Stylesheet _target;
            private readonly WorkbookStyleColorResolver _sourceColors;
            private readonly WorkbookStyleColorResolver _targetColors;
            private readonly NumberFormatRegistry _numberFormats;

            internal StylesheetMerger(WorkbookPart sourceWorkbookPart, Stylesheet source, WorkbookPart targetWorkbookPart, Stylesheet target) {
                _source = source;
                _target = target;
                _sourceColors = WorkbookStyleColorResolver.Create(sourceWorkbookPart, source);
                _targetColors = WorkbookStyleColorResolver.Create(targetWorkbookPart, target);
                _numberFormats = NumberFormatRegistry.Create(target);
            }

            internal WorksheetStyleCopyMap Merge() {
                Dictionary<uint, uint> numberFormatMap = MergeNumberingFormats();
                StyleEntryIndex fonts = IndexEntries<Font>(_target.Fonts!, reservesFills: false);
                Dictionary<uint, uint> fontMap = MergeEntries<Font>(_source.Fonts, _target.Fonts!, fonts);
                StyleEntryIndex fills = IndexEntries<Fill>(_target.Fills!, reservesFills: true);
                Dictionary<uint, uint> fillMap = MergeEntries<Fill>(_source.Fills, _target.Fills!, fills);
                StyleEntryIndex borders = IndexEntries<Border>(_target.Borders!, reservesFills: false);
                Dictionary<uint, uint> borderMap = MergeEntries<Border>(_source.Borders, _target.Borders!, borders);

                var cellReferences = new CellFormatReferences(_numberFormats, numberFormatMap, fonts, fontMap, fills, fillMap, borders, borderMap);
                StyleEntryIndex styleFormats = IndexCellFormats(_target.CellStyleFormats!, cellReferences, styleFormats: null);
                Dictionary<uint, uint> styleFormatMap = MergeCellFormats(_source.CellStyleFormats, _target.CellStyleFormats!, styleFormats, cellReferences, styleFormats: null, styleFormatMap: null);
                StyleEntryIndex cellFormats = IndexCellFormats(_target.CellFormats!, cellReferences, styleFormats);
                Dictionary<uint, uint> cellFormatMap = MergeCellFormats(_source.CellFormats, _target.CellFormats!, cellFormats, cellReferences, styleFormats, styleFormatMap);

                StyleEntryIndex differentialFormats = IndexEntries<DifferentialFormat>(_target.DifferentialFormats!, reservesFills: false);
                Dictionary<uint, uint> differentialFormatMap = MergeEntries<DifferentialFormat>(
                    _source.DifferentialFormats,
                    _target.DifferentialFormats!,
                    differentialFormats,
                    RemapDifferentialNumberFormats);
                return new WorksheetStyleCopyMap(cellFormatMap, differentialFormatMap);
            }

            private Dictionary<uint, uint> MergeNumberingFormats() {
                var map = new Dictionary<uint, uint>();
                if (_source.NumberingFormats == null) {
                    return map;
                }

                foreach (NumberingFormat sourceFormat in _source.NumberingFormats.Elements<NumberingFormat>()) {
                    uint? sourceId = ReadCopiedStyleIndex(sourceFormat.NumberFormatId);
                    if (!sourceId.HasValue || sourceId.Value < FirstCustomNumberFormatId || sourceFormat.FormatCode?.Value == null) {
                        continue;
                    }

                    map[sourceId.Value] = _numberFormats.GetOrList(sourceFormat);
                }

                return map;
            }

            private void RemapDifferentialNumberFormats(OpenXmlElement differentialFormat) {
                foreach (NumberingFormat format in differentialFormat.Elements<NumberingFormat>()) {
                    uint? id = ReadCopiedStyleIndex(format.NumberFormatId);
                    string? code = format.FormatCode?.Value;
                    if (id.HasValue && id.Value >= FirstCustomNumberFormatId && code != null) {
                        format.NumberFormatId = _numberFormats.GetOrReserve(code);
                    }
                }
            }

            /// <summary>
            /// Indexes target fonts, fills, borders and dxfs by their stored properties. A theme or palette reference must
            /// not alias a fixed RGB color: later theme or palette edits change only the referenced color.
            /// </summary>
            private StyleEntryIndex IndexEntries<TElement>(OpenXmlCompositeElement target, bool reservesFills)
                where TElement : OpenXmlElement {
                var index = new StyleEntryIndex();
                foreach (TElement entry in target.Elements<TElement>()) {
                    string key = BuildEntryKey(entry);
                    if (reservesFills && index.Count < 2U && !IsReservedFillKey(index.Count, key)) {
                        // Excel draws the first two fills as none and gray125 whatever they hold.
                        index.AddUnmatched();
                        continue;
                    }

                    index.Add(key);
                }

                return index;
            }

            private Dictionary<uint, uint> MergeEntries<TElement>(
                OpenXmlCompositeElement? source,
                OpenXmlCompositeElement target,
                StyleEntryIndex targetIndex,
                Action<OpenXmlElement>? prepare = null)
                where TElement : OpenXmlElement {
                var map = new Dictionary<uint, uint>();
                if (source == null) {
                    return map;
                }

                uint sourceIndex = 0;
                foreach (TElement entry in source.Elements<TElement>()) {
                    OpenXmlElement copied = entry.CloneNode(true);
                    prepare?.Invoke(copied);
                    // A copy is appended with fixed colors, but may reuse a target entry that keeps the source's theme
                    // and palette references, where the target resolves them to the same colors.
                    OpenXmlElement? bound = _sourceColors.CloneKeepingReferencesResolvedAlike(copied, _targetColors);
                    _sourceColors.NormalizeThemeAndIndexedColors(copied);
                    string? boundKey = null;
                    if (bound != null) {
                        boundKey = BuildEntryKey(bound);
                    }

                    map[sourceIndex] = AppendOrReuse(target, targetIndex, copied, BuildEntryKey(copied), boundKey);
                    sourceIndex++;
                }

                return map;
            }

            private static StyleEntryIndex IndexCellFormats(OpenXmlCompositeElement target, CellFormatReferences references, StyleEntryIndex? styleFormats) {
                var index = new StyleEntryIndex();
                foreach (CellFormat format in target.Elements<CellFormat>()) {
                    index.Add(references.BuildKey(format, styleFormats));
                }

                return index;
            }

            private static Dictionary<uint, uint> MergeCellFormats(
                OpenXmlCompositeElement? source,
                OpenXmlCompositeElement target,
                StyleEntryIndex targetIndex,
                CellFormatReferences references,
                StyleEntryIndex? styleFormats,
                IReadOnlyDictionary<uint, uint>? styleFormatMap) {
                var map = new Dictionary<uint, uint>();
                if (source == null) {
                    return map;
                }

                uint sourceIndex = 0;
                foreach (CellFormat sourceFormat in source.Elements<CellFormat>()) {
                    var copied = (CellFormat)sourceFormat.CloneNode(true);
                    references.Remap(copied);
                    if (styleFormatMap != null) {
                        copied.FormatId = MapStyleReference(copied.FormatId, styleFormatMap);
                    }

                    map[sourceIndex] = AppendOrReuse(target, targetIndex, copied, references.BuildKey(copied, styleFormats));
                    sourceIndex++;
                }

                return map;
            }

            /// <summary>
            /// Returns the first target entry with either key, or appends <paramref name="entry"/>, which has
            /// <paramref name="key"/>.
            /// </summary>
            private static uint AppendOrReuse(OpenXmlCompositeElement target, StyleEntryIndex targetIndex, OpenXmlElement entry, string key, string? boundKey = null) {
                bool found = targetIndex.TryFind(key, out uint existing);
                if (boundKey != null && targetIndex.TryFind(boundKey, out uint bound) && (!found || bound < existing)) {
                    existing = bound;
                    found = true;
                }

                if (found) {
                    return existing;
                }

                target.Append(entry);
                return targetIndex.Add(key);
            }

            private static bool IsReservedFillKey(uint position, string key) {
                PatternValues pattern = PatternValues.None;
                if (position == 1U) {
                    pattern = PatternValues.Gray125;
                }

                return string.Equals(key, BuildStyleKey(new Fill(new PatternFill { PatternType = pattern })), StringComparison.Ordinal);
            }
        }

        /// <summary>
        /// Maps a copied cell format's number format, font, fill and border onto the target, and keys cell formats by the
        /// first target entry each of those references renders as.
        /// </summary>
        private sealed class CellFormatReferences {
            private readonly NumberFormatRegistry _numberFormats;
            private readonly IReadOnlyDictionary<uint, uint> _numberFormatMap;
            private readonly StyleEntryIndex _fonts;
            private readonly IReadOnlyDictionary<uint, uint> _fontMap;
            private readonly StyleEntryIndex _fills;
            private readonly IReadOnlyDictionary<uint, uint> _fillMap;
            private readonly StyleEntryIndex _borders;
            private readonly IReadOnlyDictionary<uint, uint> _borderMap;

            internal CellFormatReferences(
                NumberFormatRegistry numberFormats,
                IReadOnlyDictionary<uint, uint> numberFormatMap,
                StyleEntryIndex fonts,
                IReadOnlyDictionary<uint, uint> fontMap,
                StyleEntryIndex fills,
                IReadOnlyDictionary<uint, uint> fillMap,
                StyleEntryIndex borders,
                IReadOnlyDictionary<uint, uint> borderMap) {
                _numberFormats = numberFormats;
                _numberFormatMap = numberFormatMap;
                _fonts = fonts;
                _fontMap = fontMap;
                _fills = fills;
                _fillMap = fillMap;
                _borders = borders;
                _borderMap = borderMap;
            }

            internal void Remap(CellFormat format) {
                format.NumberFormatId = MapStyleReference(format.NumberFormatId, _numberFormatMap);
                format.FontId = MapStyleReference(format.FontId, _fontMap);
                format.FillId = MapStyleReference(format.FillId, _fillMap);
                format.BorderId = MapStyleReference(format.BorderId, _borderMap);
            }

            /// <summary>
            /// Builds a key that is equal for two cell formats with the same referenced style identities. A cell style
            /// format has no parent format, so <paramref name="styleFormats"/> is null for those.
            /// </summary>
            internal string BuildKey(CellFormat format, StyleEntryIndex? styleFormats) {
                var builder = new StringBuilder();
                AppendReference(builder, format.NumberFormatId, _numberFormats.Canonical);
                AppendReference(builder, format.FontId, _fonts.Canonical);
                AppendReference(builder, format.FillId, _fills.Canonical);
                AppendReference(builder, format.BorderId, _borders.Canonical);
                if (styleFormats != null) {
                    AppendReference(builder, format.FormatId, styleFormats.Canonical);
                }

                AppendStyleKey(builder, format, CellFormatReferenceAttributes);
                return builder.ToString();
            }

            private static void AppendReference(StringBuilder builder, UInt32Value? value, Func<uint, uint> canonical) {
                if (value == null) {
                    builder.Append(canonical(0U).ToString(CultureInfo.InvariantCulture));
                } else if (ReadCopiedStyleIndex(value) is uint id) {
                    builder.Append(canonical(id).ToString(CultureInfo.InvariantCulture));
                } else {
                    // Unreadable references never equal another format.
                    AppendKeyText(builder.Append('?'), value.InnerText);
                }

                builder.Append(';');
            }
        }

        /// <summary>
        /// Finds the first target style entry with a given key. Each target entry also records the first identical earlier
        /// entry, so cell formats can compare equivalent references without losing their color bindings.
        /// </summary>
        private sealed class StyleEntryIndex {
            private readonly Dictionary<string, uint> _firstIndexes = new Dictionary<string, uint>(StringComparer.Ordinal);
            private readonly List<uint> _canonicalIndexes = new List<uint>();

            internal uint Count => (uint)_canonicalIndexes.Count;

            /// <summary>Records the next entry and returns the index of the first entry with the same key.</summary>
            internal uint Add(string key) {
                uint canonical = Count;
                if (_firstIndexes.TryGetValue(key, out uint first)) {
                    canonical = Math.Min(canonical, first);
                }

                AddFirst(key, canonical);

                _canonicalIndexes.Add(canonical);
                return canonical;
            }

            /// <summary>Records the next entry as one no copied entry may reuse.</summary>
            internal void AddUnmatched() => _canonicalIndexes.Add(Count);

            internal bool TryFind(string key, out uint index) => _firstIndexes.TryGetValue(key, out index);

            internal uint Canonical(uint index) {
                if (index < Count) {
                    return _canonicalIndexes[(int)index];
                }

                return index;
            }

            private void AddFirst(string key, uint index) {
                if (!_firstIndexes.ContainsKey(key)) {
                    _firstIndexes.Add(key, index);
                }
            }
        }

        /// <summary>
        /// Keeps each custom number format id naming one format code across the target's numFmts and dxfs. Excel reads both
        /// into one table, so a copied format must neither take an id the target uses for another code nor add a second id
        /// for a code the target already has.
        /// </summary>
        private sealed class NumberFormatRegistry {
            private readonly Stylesheet _stylesheet;
            private readonly Dictionary<string, uint> _idsByCode = new Dictionary<string, uint>(StringComparer.Ordinal);
            private readonly Dictionary<uint, string> _codesById = new Dictionary<uint, string>();
            private readonly Dictionary<uint, uint> _canonicalIds = new Dictionary<uint, uint>();
            private readonly HashSet<string> _listedCodes = new HashSet<string>(StringComparer.Ordinal);
            private uint _nextId = FirstCustomNumberFormatId;

            private NumberFormatRegistry(Stylesheet stylesheet) {
                _stylesheet = stylesheet;
            }

            internal static NumberFormatRegistry Create(Stylesheet stylesheet) {
                var registry = new NumberFormatRegistry(stylesheet);
                if (stylesheet.NumberingFormats != null) {
                    foreach (NumberingFormat format in stylesheet.NumberingFormats.Elements<NumberingFormat>()) {
                        registry.Register(format, listed: true);
                    }
                }

                if (stylesheet.DifferentialFormats != null) {
                    foreach (NumberingFormat format in stylesheet.DifferentialFormats.Elements<DifferentialFormat>().SelectMany(_ => _.Elements<NumberingFormat>())) {
                        registry.Register(format, listed: false);
                    }
                }

                return registry;
            }

            /// <summary>Maps a target id to the first id the target uses for the same code.</summary>
            internal uint Canonical(uint id) {
                if (_canonicalIds.TryGetValue(id, out uint canonical)) {
                    return canonical;
                }

                return id;
            }

            /// <summary>Returns the target id for a source numFmts entry, and lists a copy of it when the target does not list its code.</summary>
            internal uint GetOrList(NumberingFormat sourceFormat) {
                string code = sourceFormat.FormatCode!.Value!;
                uint id = GetOrReserve(code);
                if (_listedCodes.Add(code)) {
                    var copied = (NumberingFormat)sourceFormat.CloneNode(true);
                    copied.NumberFormatId = id;
                    NumberingFormats formats = _stylesheet.NumberingFormats ?? InsertNumberingFormats(_stylesheet);
                    formats.Append(copied);
                    formats.Count = (uint)formats.Elements<NumberingFormat>().Count();
                }

                return id;
            }

            /// <summary>Returns the target id for a code, reserving an unused id when the target has no format with that code.</summary>
            internal uint GetOrReserve(string code) {
                if (_idsByCode.TryGetValue(code, out uint id)) {
                    return id;
                }

                id = _nextId;
                _nextId = checked(_nextId + 1U);
                _idsByCode.Add(code, id);
                _codesById.Add(id, code);
                _canonicalIds.Add(id, id);
                return id;
            }

            private void Register(NumberingFormat format, bool listed) {
                uint? id = ReadCopiedStyleIndex(format.NumberFormatId);
                string? code = format.FormatCode?.Value;
                if (!id.HasValue || id.Value < FirstCustomNumberFormatId || code == null) {
                    return;
                }

                if (id.Value >= _nextId && id.Value < uint.MaxValue) {
                    _nextId = id.Value + 1U;
                }

                if (_codesById.TryGetValue(id.Value, out string? existingCode)) {
                    // An id the target already uses for another code cannot also stand for this one.
                    if (!string.Equals(existingCode, code, StringComparison.Ordinal)) {
                        return;
                    }
                } else {
                    _codesById.Add(id.Value, code);
                }

                if (!_idsByCode.TryGetValue(code, out uint first)) {
                    first = id.Value;
                    _idsByCode.Add(code, first);
                }

                if (!_canonicalIds.ContainsKey(id.Value)) {
                    _canonicalIds.Add(id.Value, first);
                }

                if (listed) {
                    _listedCodes.Add(code);
                }
            }
        }

        private static uint MapStyleReference(UInt32Value? value, IReadOnlyDictionary<uint, uint> map) {
            // A missing reference means entry 0, which maps like any other.
            uint id = ReadCopiedStyleIndex(value) ?? 0U;
            if (map.TryGetValue(id, out uint mapped)) {
                return mapped;
            }

            return id;
        }

        private static readonly string NoneFillKey = BuildStyleKey(new Fill(new PatternFill { PatternType = PatternValues.None }));

        /// <summary>
        /// Builds the key a font, fill, border or dxf is matched by. Forms that are written differently but draw nothing
        /// different share a key: a border side with no line style is left out, so Excel's default border with its empty
        /// sides matches an empty border, and a fill with no pattern matches the none pattern. A dxf is keyed as written:
        /// it overrides only what it names, so these forms are not taken as equal inside one.
        /// </summary>
        private static string BuildEntryKey(OpenXmlElement entry) {
            if (entry is Fill fill && !fill.HasChildren && !fill.HasAttributes) {
                return NoneFillKey;
            }

            if (entry is Border border && border.ChildElements.Any(IsEmptyBorderSide)) {
                var drawn = (Border)border.CloneNode(true);
                foreach (OpenXmlElement side in drawn.ChildElements.Where(IsEmptyBorderSide).ToList()) {
                    side.Remove();
                }

                return BuildStyleKey(drawn);
            }

            return BuildStyleKey(entry);
        }

        private static bool IsEmptyBorderSide(OpenXmlElement side) {
            if (side is not BorderPropertiesType || side.HasChildren) {
                return false;
            }

            // A side without a style draws no line, as does the none style.
            return side.GetAttributes().All(_ => string.IsNullOrEmpty(_.NamespaceUri) && _.LocalName == "style" && _.Value == "none");
        }

        /// <summary>
        /// Builds a key that is equal for two style entries exactly when they hold the same elements, attributes and text,
        /// whatever namespace prefixes their packages use.
        /// </summary>
        private static string BuildStyleKey(OpenXmlElement element) {
            var builder = new StringBuilder();
            AppendStyleKey(builder, element, ignoredAttributes: null);
            return builder.ToString();
        }

        private static void AppendStyleKey(StringBuilder builder, OpenXmlElement element, ICollection<string>? ignoredAttributes) {
            builder.Append('<').Append(element.NamespaceUri).Append('|').Append(element.LocalName);
            if (element.HasAttributes) {
                IEnumerable<OpenXmlAttribute> attributes = element.GetAttributes()
                    .Where(_ => ignoredAttributes == null || !string.IsNullOrEmpty(_.NamespaceUri) || !ignoredAttributes.Contains(_.LocalName))
                    .OrderBy(_ => _.NamespaceUri, StringComparer.Ordinal)
                    .ThenBy(_ => _.LocalName, StringComparer.Ordinal);
                foreach (OpenXmlAttribute attribute in attributes) {
                    builder.Append(' ').Append(attribute.NamespaceUri).Append('|').Append(attribute.LocalName).Append('=');
                    AppendKeyText(builder, attribute.Value);
                }
            }

            builder.Append('>');
            if (element.HasChildren) {
                foreach (OpenXmlElement child in element.ChildElements) {
                    AppendStyleKey(builder, child, ignoredAttributes: null);
                }
            } else {
                AppendKeyText(builder, element.InnerText);
            }

            builder.Append("</>");
        }

        // Length-prefixed, so no value can run into the text that follows it.
        private static void AppendKeyText(StringBuilder builder, string? text) {
            text ??= string.Empty;
            builder.Append(text.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(text);
        }
    }
}

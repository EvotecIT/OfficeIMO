using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private const uint FirstCustomNumberFormatId = 164U;

        // The attributes a cell format key reads through the target's indexes rather than as text.
        private static readonly string[] CellFormatReferenceAttributes = { "numFmtId", "fontId", "fillId", "borderId", "xfId" };

        /// <summary>
        /// Merges a source workbook's styles into the target workbook for one worksheet copy. Each source entry maps to the
        /// first target entry that renders the same, and is appended only when the target has none, so copying more sheets
        /// from the same source, or from a source with the same styles, adds nothing. The target is indexed afresh for every
        /// copy, so changes to either stylesheet between copies are always seen, and entries are visited in index order,
        /// so equal inputs give equal output.
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
            /// Indexes target fonts, fills, borders or dxfs both by their own XML and with their colors resolved through the
            /// target's theme and palette, so a copied entry whose resolved colors match reuses the target's entry.
            /// </summary>
            private StyleEntryIndex IndexEntries<TElement>(OpenXmlCompositeElement target, bool reservesFills)
                where TElement : OpenXmlElement {
                var index = new StyleEntryIndex();
                foreach (TElement entry in target.Elements<TElement>()) {
                    string key = BuildStyleKey(entry);
                    if (reservesFills && index.Count < 2U && !IsReservedFillKey(index.Count, key)) {
                        // Excel draws the first two fills as none and gray125 whatever they hold.
                        index.AddUnmatched();
                        continue;
                    }

                    index.Add(key, BuildResolvedColorKey(entry));
                }

                return index;
            }

            private string? BuildResolvedColorKey(OpenXmlElement entry) {
                if (!entry.Descendants<ColorType>().Any()) {
                    return null;
                }

                OpenXmlElement resolved = entry.CloneNode(true);
                _targetColors.NormalizeThemeAndIndexedColors(resolved);
                return BuildStyleKey(resolved);
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
                    _sourceColors.NormalizeThemeAndIndexedColors(copied);
                    prepare?.Invoke(copied);
                    map[sourceIndex] = AppendOrReuse(target, targetIndex, copied, BuildStyleKey(copied));
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

            private static uint AppendOrReuse(OpenXmlCompositeElement target, StyleEntryIndex targetIndex, OpenXmlElement entry, string key) {
                if (targetIndex.TryFind(key, out uint existing)) {
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
            /// Builds a key that is equal for two cell formats exactly when they render the same in the target. A cell style
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
        /// Finds the first target style entry with a given key. Each target entry also records the first earlier entry that
        /// renders the same, so cell formats can be compared by what their fonts, fills and borders render as.
        /// </summary>
        private sealed class StyleEntryIndex {
            private readonly Dictionary<string, uint> _firstIndexes = new Dictionary<string, uint>(StringComparer.Ordinal);
            private readonly List<uint> _canonicalIndexes = new List<uint>();

            internal uint Count => (uint)_canonicalIndexes.Count;

            /// <summary>Records the next entry and returns the index of the first entry that renders the same.</summary>
            internal uint Add(string key, string? equivalentKey = null) {
                uint canonical = Count;
                if (_firstIndexes.TryGetValue(key, out uint first)) {
                    canonical = Math.Min(canonical, first);
                }

                if (equivalentKey != null && _firstIndexes.TryGetValue(equivalentKey, out first)) {
                    canonical = Math.Min(canonical, first);
                }

                AddFirst(key, canonical);
                if (equivalentKey != null) {
                    AddFirst(equivalentKey, canonical);
                }

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

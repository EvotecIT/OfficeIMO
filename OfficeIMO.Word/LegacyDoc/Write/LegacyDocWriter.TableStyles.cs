using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        /// <summary>
        /// The table style a table resolves to: its own <c>w:tblStyle</c> value, or the document's default table style
        /// when it names none. Readers see the default as if the table had named it.
        /// </summary>
        private static string? ResolveSupportedEffectiveTableStyleId(LegacyDocElement properties,
            LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            string? styleId = properties.Child("tblStyle").Val;
            if (!string.IsNullOrWhiteSpace(styleId)) {
                return styleId;
            }
            LegacyDocElement defaultStyle = tableStyleDefinitions.DefaultStyle;
            if (defaultStyle.IsNone || IsNoOpTableStyle(LegacyDocStyleCatalog.StyleId(defaultStyle))) {
                return styleId;
            }
            // Resolve an implicit document default without adding a style to
            // the caller's source XML. Existing style readers validate it.
            return LegacyDocStyleCatalog.StyleId(defaultStyle)!;
        }

        private static LegacyDocTableStyleDefinitions ReadTableStyleDefinitions(LegacyDocSourcePackage source) =>
            new(source.Styles);

        /// <summary>
        /// The document's table styles by identifier, first definition winning, read from its style catalog. A style
        /// is parsed only when a table looks it up.
        /// </summary>
        internal sealed class LegacyDocTableStyleDefinitions : IReadOnlyDictionary<string, LegacyDocElement> {
            private readonly Dictionary<string, LegacyDocElement> _styles = new(StringComparer.Ordinal);
            private readonly LegacyDocElement _defaultStyle;

            internal LegacyDocTableStyleDefinitions(LegacyDocStyleCatalog catalog) {
                foreach (LegacyDocElement style in catalog.Definitions) {
                    string? styleId = LegacyDocStyleCatalog.StyleId(style);
                    if (!string.Equals(LegacyDocStyleCatalog.Type(style), "table", StringComparison.Ordinal) || string.IsNullOrWhiteSpace(styleId)) continue;
                    if (_styles.ContainsKey(styleId!)) continue;
                    _styles.Add(styleId!, style);
                    if (_defaultStyle.IsNone && IsDefaultTableStyle(style)) _defaultStyle = style;
                }
            }

            // The style catalog reads w:default as the SDK's OnOffValue.HasValue did: a value it can't parse is not on.
            private static bool IsDefaultTableStyle(LegacyDocElement style) =>
                style.Attr("default") is "true" or "1" or "on";

            /// <summary>The first table style, in document order, that is marked as the default.</summary>
            internal LegacyDocElement DefaultStyle => _defaultStyle;

            public LegacyDocElement this[string key] => _styles[key];
            public IEnumerable<string> Keys => _styles.Keys;
            public IEnumerable<LegacyDocElement> Values => _styles.Values;
            public int Count => _styles.Count;
            public bool ContainsKey(string key) => _styles.ContainsKey(key);

            public bool TryGetValue(string key, out LegacyDocElement value) => _styles.TryGetValue(key, out value);

            public IEnumerator<KeyValuePair<string, LegacyDocElement>> GetEnumerator() => _styles.GetEnumerator();

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }

        private static LegacyDocTableBorders ReadSupportedTableStyleBorders(string? styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            if (IsNoOpTableStyle(styleId)) {
                return default;
            }

            if (string.IsNullOrWhiteSpace(styleId) || !TryResolveTableStyleDefinition(styleId!, tableStyleDefinitions, out LegacyDocElement style)) {
                throw new NotSupportedException($"Native DOC saving supports simple tables only when table style '{styleId}' can be resolved to supported table-level formatting.");
            }

            ThrowIfUnsupportedTableStyle(styleId!, style, tableStyleDefinitions);
            LegacyDocTableBorders inheritedBorders = ReadSupportedTableStyleBaseBorders(style, tableStyleDefinitions);
            LegacyDocElement customBorders = style.Child("tblPr").Child("tblBorders");
            LegacyDocTableBorders ownBorders = customBorders.IsNone ? default : ReadSupportedTableBorders(customBorders);
            return MergeSupportedTableBorders(ownBorders, inheritedBorders);
        }

        private static LegacyDocTableCellShading ReadSupportedTableStyleShading(string? styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            if (IsNoOpTableStyle(styleId)) {
                return default;
            }

            if (string.IsNullOrWhiteSpace(styleId) || !TryResolveTableStyleDefinition(styleId!, tableStyleDefinitions, out LegacyDocElement style)) {
                throw new NotSupportedException($"Native DOC saving supports simple tables only when table style '{styleId}' can be resolved to supported table-level formatting.");
            }

            ThrowIfUnsupportedTableStyle(styleId!, style, tableStyleDefinitions);
            LegacyDocTableCellShading inheritedShading = ReadSupportedTableStyleBaseShading(style, tableStyleDefinitions);
            LegacyDocTableCellShading ownShading = ReadSupportedTableStyleOwnShading(style);
            return ownShading.IsSpecified ? ownShading : inheritedShading;
        }

        private static LegacyDocTableCellMargins? ReadSupportedTableStyleDefaultCellMargins(string? styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            LegacyDocElement style = ResolveSupportedTableStyle(styleId, tableStyleDefinitions);
            if (style.IsNone) {
                return null;
            }

            LegacyDocTableCellMargins? inheritedMargins = ReadSupportedTableStyleBaseDefaultCellMargins(style, tableStyleDefinitions);
            LegacyDocTableCellMargins? ownMargins = ReadSupportedTableStyleOwnDefaultCellMargins(style);
            if (ownMargins == null) {
                return inheritedMargins;
            }

            return inheritedMargins == null
                ? ownMargins
                : inheritedMargins.Value.Merge(ownMargins.Value);
        }

        private static int? ReadSupportedTableStyleDefaultCellSpacing(string? styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            LegacyDocElement style = ResolveSupportedTableStyle(styleId, tableStyleDefinitions);
            return style.IsNone
                ? null
                : ReadSupportedTableStyleOwnDefaultCellSpacing(style) ?? ReadSupportedTableStyleBaseDefaultCellSpacing(style, tableStyleDefinitions);
        }

        private static LegacyDocTableAlignment? ReadSupportedTableStyleAlignment(string? styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            LegacyDocElement style = ResolveSupportedTableStyle(styleId, tableStyleDefinitions);
            return style.IsNone
                ? null
                : ReadSupportedTableStyleOwnAlignment(style) ?? ReadSupportedTableStyleBaseAlignment(style, tableStyleDefinitions);
        }

        private static int? ReadSupportedTableStyleIndentation(string? styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            LegacyDocElement style = ResolveSupportedTableStyle(styleId, tableStyleDefinitions);
            return style.IsNone
                ? null
                : ReadSupportedTableStyleOwnIndentation(style) ?? ReadSupportedTableStyleBaseIndentation(style, tableStyleDefinitions);
        }

        private static LegacyDocTablePreferredWidth? ReadSupportedTableStylePreferredWidth(string? styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            LegacyDocElement style = ResolveSupportedTableStyle(styleId, tableStyleDefinitions);
            return style.IsNone
                ? null
                : ReadSupportedTableStyleOwnPreferredWidth(style) ?? ReadSupportedTableStyleBasePreferredWidth(style, tableStyleDefinitions);
        }

        private static bool? ReadSupportedTableStyleAutofit(string? styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            LegacyDocElement style = ResolveSupportedTableStyle(styleId, tableStyleDefinitions);
            return style.IsNone
                ? null
                : ReadSupportedTableStyleOwnAutofit(style) ?? ReadSupportedTableStyleBaseAutofit(style, tableStyleDefinitions);
        }

        private static LegacyDocWritableParagraphFormatting ReadSupportedTableStyleParagraphFormatting(string? styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            if (IsNoOpTableStyle(styleId)) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            LegacyDocElement style = ResolveSupportedTableStyle(styleId, tableStyleDefinitions);
            if (style.IsNone) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            LegacyDocWritableParagraphFormatting inheritedFormatting = ReadSupportedTableStyleBaseParagraphFormatting(style, tableStyleDefinitions);
            LegacyDocWritableParagraphFormatting ownFormatting = ReadSupportedTableStyleOwnParagraphFormatting(style);
            return ownFormatting.WithInheritedParagraphFormatting(inheritedFormatting);
        }

        private static LegacyDocWritableFormatting ReadSupportedTableStyleRunFormatting(string? styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            if (IsNoOpTableStyle(styleId)) {
                return LegacyDocWritableFormatting.Plain;
            }

            LegacyDocElement style = ResolveSupportedTableStyle(styleId, tableStyleDefinitions);
            if (style.IsNone) {
                return LegacyDocWritableFormatting.Plain;
            }

            LegacyDocWritableFormatting inheritedFormatting = ReadSupportedTableStyleBaseRunFormatting(style, tableStyleDefinitions);
            LegacyDocWritableFormatting ownFormatting = ReadSupportedTableStyleOwnRunFormatting(style);
            return ownFormatting.WithInheritedFormatting(inheritedFormatting);
        }

        private static LegacyDocElement ResolveSupportedTableStyle(string? styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            if (IsNoOpTableStyle(styleId)) {
                return LegacyDocElement.None;
            }

            if (string.IsNullOrWhiteSpace(styleId) || !TryResolveTableStyleDefinition(styleId!, tableStyleDefinitions, out LegacyDocElement style)) {
                throw new NotSupportedException($"Native DOC saving supports simple tables only when table style '{styleId}' can be resolved to supported table-level formatting.");
            }

            ThrowIfUnsupportedTableStyle(styleId!, style, tableStyleDefinitions);
            return style;
        }

        private static bool TryResolveTableStyleDefinition(string styleId, LegacyDocTableStyleDefinitions tableStyleDefinitions, out LegacyDocElement style) {
            // The document owns authored built-in definitions; library defaults fill only a missing definition.
            if (tableStyleDefinitions.TryGetValue(styleId, out style)) {
                return true;
            }
            if (IsTableGridStyle(styleId)) {
                style = LegacyDocWordHelpers.BuiltInTableStyle(WordTableStyle.TableGrid);
                return true;
            }
            style = LegacyDocElement.None;
            return false;
        }

        private static bool IsNoOpTableStyle(string? styleId) {
            if (string.IsNullOrWhiteSpace(styleId)) {
                return true;
            }

            // Word ignores TableNormal child properties, including when it is a style base.
            return string.Equals(styleId, "TableNormal", StringComparison.Ordinal)
                || string.Equals(styleId, "NormalTable", StringComparison.Ordinal);
        }

        private static bool IsTableGridStyle(string? styleId) =>
            string.Equals(styleId, "TableGrid", StringComparison.Ordinal);

        private static void ThrowIfUnsupportedTableStyle(string styleId, LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            if (LegacyDocStyleCatalog.Type(style) != "table") {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' only when it is a table style.");
            }

            foreach (LegacyDocElement child in style.Elements()) {
                switch (child.WName) {
                    case "name":
                    case "uiPriority":
                    case "rsid":
                    case "semiHidden":
                    case "unhideWhenUsed":
                    case "qFormat":
                    case "locked":
                    case "stylePaneFormatFilter":
                        break;
                    case "basedOn":
                        ThrowIfUnsupportedTableStyleBase(styleId, child, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal) { styleId });
                        break;
                    case "tblPr":
                        ThrowIfUnsupportedStyleTableProperties(styleId, child);
                        break;
                    case "tcPr":
                        ThrowIfUnsupportedStyleTableCellProperties(styleId, child);
                        break;
                    case "tblStylePr":
                        ThrowIfUnsupportedTableStyleConditionalProperties(styleId, child);
                        break;
                    case "pPr":
                        ThrowIfUnsupportedTableStyleParagraphProperties(styleId, child);
                        break;
                    case "rPr":
                        ThrowIfUnsupportedTableStyleRunProperties(styleId, child);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving does not support table style '{styleId}' element '{child.LocalName}'.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableStyleParagraphProperties(string styleId, LegacyDocElement paragraphProperties) {
            try {
                _ = ReadSupportedStyleParagraphFormatting(paragraphProperties);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' paragraph formatting only with supported paragraph properties. {exception.Message}", exception);
            }
        }

        private static void ThrowIfUnsupportedTableStyleRunProperties(string styleId, LegacyDocElement runProperties) {
            try {
                _ = ReadSupportedRunFormatting(runProperties);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' run formatting only with supported run properties. {exception.Message}", exception);
            }
        }

        private static void ThrowIfUnsupportedTableStyleBase(string styleId, LegacyDocElement basedOn, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            string? baseStyleId = basedOn.Val;
            if (IsNoOpTableStyle(baseStyleId)) {
                return;
            }

            if (string.IsNullOrWhiteSpace(baseStyleId)
                || !TryResolveTableStyleDefinition(baseStyleId!, tableStyleDefinitions, out LegacyDocElement baseStyle)) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' only when its base style can be resolved to supported table-level formatting.");
            }

            ThrowIfUnsupportedInheritedTableStyleBase(styleId, baseStyleId!, baseStyle, tableStyleDefinitions, visitedStyleIds);
        }

        private static LegacyDocTableBorders ReadSupportedTableStyleBaseBorders(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseBorders(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static LegacyDocTableCellShading ReadSupportedTableStyleBaseShading(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseShading(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static LegacyDocTableCellMargins? ReadSupportedTableStyleBaseDefaultCellMargins(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseDefaultCellMargins(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static int? ReadSupportedTableStyleBaseDefaultCellSpacing(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseDefaultCellSpacing(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static LegacyDocTableAlignment? ReadSupportedTableStyleBaseAlignment(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseAlignment(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static int? ReadSupportedTableStyleBaseIndentation(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseIndentation(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static LegacyDocTablePreferredWidth? ReadSupportedTableStyleBasePreferredWidth(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBasePreferredWidth(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static bool? ReadSupportedTableStyleBaseAutofit(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseAutofit(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static int? ReadSupportedTableStyleBaseRowBandSize(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseRowBandSize(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static int? ReadSupportedTableStyleBaseColumnBandSize(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseColumnBandSize(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static LegacyDocWritableParagraphFormatting ReadSupportedTableStyleBaseParagraphFormatting(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseParagraphFormatting(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static LegacyDocWritableFormatting ReadSupportedTableStyleBaseRunFormatting(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseRunFormatting(style, tableStyleDefinitions, new HashSet<string>(StringComparer.Ordinal));
        }

        private static LegacyDocTableBorders ReadSupportedTableStyleBaseBorders(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            string? baseStyleId = style.Child("basedOn").Val;
            if (IsNoOpTableStyle(baseStyleId)) {
                return default;
            }

            if (string.IsNullOrWhiteSpace(baseStyleId)
                || !TryResolveTableStyleDefinition(baseStyleId!, tableStyleDefinitions, out LegacyDocElement baseStyle)) {
                return default;
            }

            string currentStyleId = LegacyDocStyleCatalog.StyleId(style) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(currentStyleId) && !visitedStyleIds.Add(currentStyleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write table style '{currentStyleId}' because its basedOn chain contains a cycle.");
            }

            ThrowIfUnsupportedInheritedTableStyleBase(currentStyleId, baseStyleId!, baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocTableBorders inheritedBorders = ReadSupportedTableStyleBaseBorders(baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocTableBorders baseBorders = ReadSupportedTableStyleOwnBorders(baseStyle);
            return MergeSupportedTableBorders(baseBorders, inheritedBorders);
        }

        private static LegacyDocTableCellShading ReadSupportedTableStyleBaseShading(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            string? baseStyleId = style.Child("basedOn").Val;
            if (IsNoOpTableStyle(baseStyleId)) {
                return default;
            }

            if (string.IsNullOrWhiteSpace(baseStyleId)
                || !TryResolveTableStyleDefinition(baseStyleId!, tableStyleDefinitions, out LegacyDocElement baseStyle)) {
                return default;
            }

            string currentStyleId = LegacyDocStyleCatalog.StyleId(style) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(currentStyleId) && !visitedStyleIds.Add(currentStyleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write table style '{currentStyleId}' because its basedOn chain contains a cycle.");
            }

            ThrowIfUnsupportedInheritedTableStyleBase(currentStyleId, baseStyleId!, baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocTableCellShading inheritedShading = ReadSupportedTableStyleBaseShading(baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocTableCellShading baseShading = ReadSupportedTableStyleOwnShading(baseStyle);
            return baseShading.IsSpecified ? baseShading : inheritedShading;
        }

        private static LegacyDocTableCellMargins? ReadSupportedTableStyleBaseDefaultCellMargins(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            return ReadSupportedTableStyleBaseValue(
                style,
                tableStyleDefinitions,
                visitedStyleIds,
                ReadSupportedTableStyleOwnDefaultCellMargins,
                (own, inherited) => inherited == null ? own : inherited.Value.Merge(own.GetValueOrDefault()));
        }

        private static int? ReadSupportedTableStyleBaseDefaultCellSpacing(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            return ReadSupportedTableStyleBaseValue(style, tableStyleDefinitions, visitedStyleIds, ReadSupportedTableStyleOwnDefaultCellSpacing);
        }

        private static LegacyDocTableAlignment? ReadSupportedTableStyleBaseAlignment(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            return ReadSupportedTableStyleBaseValue(style, tableStyleDefinitions, visitedStyleIds, ReadSupportedTableStyleOwnAlignment);
        }

        private static int? ReadSupportedTableStyleBaseIndentation(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            return ReadSupportedTableStyleBaseValue(style, tableStyleDefinitions, visitedStyleIds, ReadSupportedTableStyleOwnIndentation);
        }

        private static LegacyDocTablePreferredWidth? ReadSupportedTableStyleBasePreferredWidth(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            return ReadSupportedTableStyleBaseValue(style, tableStyleDefinitions, visitedStyleIds, ReadSupportedTableStyleOwnPreferredWidth);
        }

        private static bool? ReadSupportedTableStyleBaseAutofit(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            return ReadSupportedTableStyleBaseValue(style, tableStyleDefinitions, visitedStyleIds, ReadSupportedTableStyleOwnAutofit);
        }

        private static int? ReadSupportedTableStyleBaseRowBandSize(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            return ReadSupportedTableStyleBaseValue(style, tableStyleDefinitions, visitedStyleIds, ReadSupportedTableStyleOwnRowBandSize);
        }

        private static int? ReadSupportedTableStyleBaseColumnBandSize(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            return ReadSupportedTableStyleBaseValue(style, tableStyleDefinitions, visitedStyleIds, ReadSupportedTableStyleOwnColumnBandSize);
        }

        private static LegacyDocWritableParagraphFormatting ReadSupportedTableStyleBaseParagraphFormatting(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            string? baseStyleId = style.Child("basedOn").Val;
            if (IsNoOpTableStyle(baseStyleId)) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            if (string.IsNullOrWhiteSpace(baseStyleId)
                || !TryResolveTableStyleDefinition(baseStyleId!, tableStyleDefinitions, out LegacyDocElement baseStyle)) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            string currentStyleId = LegacyDocStyleCatalog.StyleId(style) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(currentStyleId) && !visitedStyleIds.Add(currentStyleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write table style '{currentStyleId}' because its basedOn chain contains a cycle.");
            }

            ThrowIfUnsupportedInheritedTableStyleBase(currentStyleId, baseStyleId!, baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocWritableParagraphFormatting inheritedFormatting = ReadSupportedTableStyleBaseParagraphFormatting(baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocWritableParagraphFormatting ownFormatting = ReadSupportedTableStyleOwnParagraphFormatting(baseStyle);
            return ownFormatting.WithInheritedParagraphFormatting(inheritedFormatting);
        }

        private static LegacyDocWritableFormatting ReadSupportedTableStyleBaseRunFormatting(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            string? baseStyleId = style.Child("basedOn").Val;
            if (IsNoOpTableStyle(baseStyleId)) {
                return LegacyDocWritableFormatting.Plain;
            }

            if (string.IsNullOrWhiteSpace(baseStyleId)
                || !TryResolveTableStyleDefinition(baseStyleId!, tableStyleDefinitions, out LegacyDocElement baseStyle)) {
                return LegacyDocWritableFormatting.Plain;
            }

            string currentStyleId = LegacyDocStyleCatalog.StyleId(style) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(currentStyleId) && !visitedStyleIds.Add(currentStyleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write table style '{currentStyleId}' because its basedOn chain contains a cycle.");
            }

            ThrowIfUnsupportedInheritedTableStyleBase(currentStyleId, baseStyleId!, baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocWritableFormatting inheritedFormatting = ReadSupportedTableStyleBaseRunFormatting(baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocWritableFormatting ownFormatting = ReadSupportedTableStyleOwnRunFormatting(baseStyle);
            return ownFormatting.WithInheritedFormatting(inheritedFormatting);
        }

        private static T? ReadSupportedTableStyleBaseValue<T>(
            LegacyDocElement style,
            LegacyDocTableStyleDefinitions tableStyleDefinitions,
            ISet<string> visitedStyleIds,
            Func<LegacyDocElement, T?> readOwnValue,
            Func<T?, T?, T?>? merge = null)
            where T : struct {
            string? baseStyleId = style.Child("basedOn").Val;
            if (IsNoOpTableStyle(baseStyleId)) {
                return null;
            }

            if (string.IsNullOrWhiteSpace(baseStyleId)
                || !TryResolveTableStyleDefinition(baseStyleId!, tableStyleDefinitions, out LegacyDocElement baseStyle)) {
                return null;
            }

            string currentStyleId = LegacyDocStyleCatalog.StyleId(style) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(currentStyleId) && !visitedStyleIds.Add(currentStyleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write table style '{currentStyleId}' because its basedOn chain contains a cycle.");
            }

            ThrowIfUnsupportedInheritedTableStyleBase(currentStyleId, baseStyleId!, baseStyle, tableStyleDefinitions, visitedStyleIds);
            T? inheritedValue = ReadSupportedTableStyleBaseValue(baseStyle, tableStyleDefinitions, visitedStyleIds, readOwnValue, merge);
            T? ownValue = readOwnValue(baseStyle);
            if (ownValue == null) {
                return inheritedValue;
            }

            return merge == null ? ownValue : merge(ownValue, inheritedValue);
        }

        private static void ThrowIfUnsupportedInheritedTableStyleBase(string styleId, string baseStyleId, LegacyDocElement baseStyle, LegacyDocTableStyleDefinitions tableStyleDefinitions, ISet<string> visitedStyleIds) {
            if (LegacyDocStyleCatalog.Type(baseStyle) != "table") {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' base style '{baseStyleId}' only when it is a table style.");
            }

            if (!visitedStyleIds.Add(baseStyleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write table style '{styleId}' because its basedOn chain contains a cycle.");
            }

            foreach (LegacyDocElement child in baseStyle.Elements()) {
                switch (child.WName) {
                    case "name":
                    case "uiPriority":
                    case "rsid":
                    case "semiHidden":
                    case "unhideWhenUsed":
                    case "qFormat":
                    case "locked":
                    case "stylePaneFormatFilter":
                        break;
                    case "basedOn":
                        ThrowIfUnsupportedTableStyleBase(baseStyleId, child, tableStyleDefinitions, visitedStyleIds);
                        break;
                    case "tblPr":
                        ThrowIfUnsupportedInheritedStyleTableProperties(styleId, baseStyleId, child);
                        break;
                    case "tcPr":
                        ThrowIfUnsupportedStyleTableCellProperties(baseStyleId, child);
                        break;
                    case "pPr":
                        ThrowIfUnsupportedTableStyleParagraphProperties(baseStyleId, child);
                        break;
                    case "rPr":
                        ThrowIfUnsupportedTableStyleRunProperties(baseStyleId, child);
                        break;
                    case "tblStylePr":
                        ThrowIfUnsupportedTableStyleConditionalProperties(baseStyleId, child);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' base style '{baseStyleId}' only when inherited custom formatting is supported table layout, borders, shading, default cell margins, default cell spacing, paragraph formatting, run formatting, and conditional formatting. Unsupported base style element: {child.LocalName}.");
                }
            }

            visitedStyleIds.Remove(baseStyleId);
        }

        private static void ThrowIfUnsupportedInheritedStyleTableProperties(string styleId, string baseStyleId, LegacyDocElement styleTableProperties) {
            foreach (LegacyDocElement child in styleTableProperties.Elements()) {
                switch (child.WName) {
                    case "tblBorders":
                        ReadSupportedTableBorders(child);
                        break;
                    case "shd":
                        ReadSupportedTableCellShading(child, "inherited table style shading");
                        break;
                    case "tblCellMar":
                        ReadSupportedTableDefaultCellMargins(child);
                        break;
                    case "tblCellSpacing":
                        ReadSupportedTableDefaultCellSpacing(child);
                        break;
                    case "jc":
                        ReadSupportedTableAlignment(child);
                        break;
                    case "tblInd":
                        ReadSupportedTableIndentation(child);
                        break;
                    case "tblW":
                        ReadSupportedTablePreferredWidth(child);
                        break;
                    case "tblLayout":
                        ReadSupportedTableAutofit(child);
                        break;
                    case "tblStyleRowBandSize":
                        ReadSupportedTableStyleBandSize(child.Val, "row");
                        break;
                    case "tblStyleColBandSize":
                        ReadSupportedTableStyleBandSize(child.Val, "column");
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' base style '{baseStyleId}' only when inherited custom table-level formatting is supported layout, borders, shading, default cell margins, and default cell spacing. Unsupported inherited table style property: {child.LocalName}.");
                }
            }
        }

        private static LegacyDocTableBorders ReadSupportedTableStyleOwnBorders(LegacyDocElement style) {
            LegacyDocElement borders = style.Child("tblPr").Child("tblBorders");
            return borders.IsNone ? default : ReadSupportedTableBorders(borders);
        }

        private static LegacyDocTableCellShading ReadSupportedTableStyleOwnShading(LegacyDocElement style) {
            // Table-level shading paints cell-spacing gaps; only tcPr shading supplies cell defaults.
            LegacyDocElement shading = style.Child("tcPr").Child("shd");
            return shading.IsNone ? default : ReadSupportedTableCellShading(shading, "table style shading");
        }

        private static LegacyDocTableCellMargins? ReadSupportedTableStyleOwnDefaultCellMargins(LegacyDocElement style) {
            LegacyDocElement margins = style.Child("tblPr").Child("tblCellMar");
            return margins.IsNone ? null : ReadSupportedTableDefaultCellMargins(margins);
        }

        private static int? ReadSupportedTableStyleOwnDefaultCellSpacing(LegacyDocElement style) {
            LegacyDocElement spacing = style.Child("tblPr").Child("tblCellSpacing");
            return spacing.IsNone ? null : ReadSupportedTableDefaultCellSpacing(spacing);
        }

        private static LegacyDocTableAlignment? ReadSupportedTableStyleOwnAlignment(LegacyDocElement style) {
            LegacyDocElement justification = style.Child("tblPr").Child("jc");
            return justification.IsNone ? null : ReadSupportedTableAlignment(justification);
        }

        private static int? ReadSupportedTableStyleOwnIndentation(LegacyDocElement style) {
            LegacyDocElement indentation = style.Child("tblPr").Child("tblInd");
            return indentation.IsNone ? null : ReadSupportedTableIndentation(indentation);
        }

        private static LegacyDocTablePreferredWidth? ReadSupportedTableStyleOwnPreferredWidth(LegacyDocElement style) {
            LegacyDocElement width = style.Child("tblPr").Child("tblW");
            return width.IsNone ? null : ReadSupportedTablePreferredWidth(width);
        }

        private static bool? ReadSupportedTableStyleOwnAutofit(LegacyDocElement style) {
            LegacyDocElement layout = style.Child("tblPr").Child("tblLayout");
            return layout.IsNone ? null : ReadSupportedTableAutofit(layout);
        }

        private static int? ReadSupportedTableStyleOwnRowBandSize(LegacyDocElement style) {
            LegacyDocElement bandSize = style.Child("tblPr").Child("tblStyleRowBandSize");
            return bandSize.IsNone ? null : ReadSupportedTableStyleBandSize(bandSize.Val, "row");
        }

        private static int? ReadSupportedTableStyleOwnColumnBandSize(LegacyDocElement style) {
            LegacyDocElement bandSize = style.Child("tblPr").Child("tblStyleColBandSize");
            return bandSize.IsNone ? null : ReadSupportedTableStyleBandSize(bandSize.Val, "column");
        }

        private static LegacyDocWritableParagraphFormatting ReadSupportedTableStyleOwnParagraphFormatting(LegacyDocElement style) {
            return ReadSupportedStyleParagraphFormatting(style.Child("pPr"));
        }

        private static LegacyDocWritableFormatting ReadSupportedTableStyleOwnRunFormatting(LegacyDocElement style) {
            return ReadSupportedRunFormatting(GetSupportedTableStyleRunProperties(style));
        }

        // A style's w:rPr, or a conditional block's: the SDK read either as the style's run properties.
        private static LegacyDocElement GetSupportedTableStyleRunProperties(LegacyDocElement styleOrConditionalProperties) {
            return styleOrConditionalProperties.Child("rPr");
        }

        private static LegacyDocTableBorders MergeSupportedTableBorders(LegacyDocTableBorders ownBorders, LegacyDocTableBorders inheritedBorders) {
            if (!inheritedBorders.HasAny || !ownBorders.HasAny) {
                return ownBorders.HasAny ? ownBorders : inheritedBorders;
            }

            return new LegacyDocTableBorders(
                ownBorders.Top.HasAny ? ownBorders.Top : inheritedBorders.Top,
                ownBorders.Left.HasAny ? ownBorders.Left : inheritedBorders.Left,
                ownBorders.Bottom.HasAny ? ownBorders.Bottom : inheritedBorders.Bottom,
                ownBorders.Right.HasAny ? ownBorders.Right : inheritedBorders.Right,
                ownBorders.InsideHorizontal.HasAny ? ownBorders.InsideHorizontal : inheritedBorders.InsideHorizontal,
                ownBorders.InsideVertical.HasAny ? ownBorders.InsideVertical : inheritedBorders.InsideVertical);
        }

        private static void ThrowIfUnsupportedStyleTableProperties(string styleId, LegacyDocElement styleTableProperties) {
            foreach (LegacyDocElement child in styleTableProperties.Elements()) {
                switch (child.WName) {
                    case "tblBorders":
                        ReadSupportedTableBorders(child);
                        break;
                    case "shd":
                        ReadSupportedTableCellShading(child, "table style shading");
                        break;
                    case "tblCellMar":
                        ReadSupportedTableDefaultCellMargins(child);
                        break;
                    case "tblCellSpacing":
                        ReadSupportedTableDefaultCellSpacing(child);
                        break;
                    case "jc":
                        ReadSupportedTableAlignment(child);
                        break;
                    case "tblInd":
                        ReadSupportedTableIndentation(child);
                        break;
                    case "tblW":
                        ReadSupportedTablePreferredWidth(child);
                        break;
                    case "tblLayout":
                        ReadSupportedTableAutofit(child);
                        break;
                    case "tblStyleRowBandSize":
                        ReadSupportedTableStyleBandSize(child.Val, "row");
                        break;
                    case "tblStyleColBandSize":
                        ReadSupportedTableStyleBandSize(child.Val, "column");
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' only with supported table-level layout, borders, shading, default cell margins, and default cell spacing. Unsupported table style property: {child.LocalName}.");
                }
            }
        }

        // A conditional region Word knows; the SDK refused any other w:type token when it read the value.
        private static void ThrowIfUnknownTableStyleConditionalType(string? styleId, string type) {
            switch (type) {
                case "wholeTable":
                case "firstRow":
                case "lastRow":
                case "firstCol":
                case "lastCol":
                case "band1Vert":
                case "band2Vert":
                case "band1Horz":
                case "band2Horz":
                case "neCell":
                case "nwCell":
                case "seCell":
                case "swCell":
                    return;
                default:
                    throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional formatting only for Word's conditional regions. Unsupported conditional type: {type}.");
            }
        }

        private static void ThrowIfUnsupportedTableStyleConditionalProperties(string styleId, LegacyDocElement tableStyleProperties) {
            string? type = tableStyleProperties.Attr("type");
            if (type == "wholeTable") return;
            if (type == null) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional formatting only when the conditional type is specified.");
            }
            ThrowIfUnknownTableStyleConditionalType(styleId, type);

            foreach (LegacyDocElement child in tableStyleProperties.Elements()) {
                switch (child.WName) {
                    case "tcPr":
                        ThrowIfUnsupportedTableStyleConditionalCellProperties(styleId, child);
                        break;
                    case "tblPr":
                        ThrowIfUnsupportedTableStyleConditionalTableProperties(styleId, child);
                        break;
                    case "trPr":
                        ThrowIfUnsupportedTableStyleConditionalRowProperties(styleId, type, child);
                        break;
                    case "pPr":
                        ThrowIfUnsupportedTableStyleConditionalParagraphProperties(styleId, child);
                        break;
                    case "rPr":
                        ThrowIfUnsupportedTableStyleConditionalRunProperties(styleId, child);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional formatting only with supported table, cell, paragraph, and run effects. Unsupported conditional style element: {child.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableStyleConditionalParagraphProperties(string styleId, LegacyDocElement paragraphProperties) {
            try {
                _ = ReadSupportedStyleParagraphFormatting(paragraphProperties);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional paragraph formatting only with supported paragraph properties. {exception.Message}", exception);
            }
        }

        private static void ThrowIfUnsupportedTableStyleConditionalRowProperties(string styleId, string type, LegacyDocElement rowProperties) {
            try {
                _ = ReadSupportedConditionalTableStyleRowProperties(rowProperties, type);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional row formatting only with supported row properties. {exception.Message}", exception);
            }
        }

        private static void ThrowIfUnsupportedTableStyleConditionalRunProperties(string styleId, LegacyDocElement runProperties) {
            try {
                _ = ReadSupportedRunFormatting(runProperties);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional run formatting only with supported run properties. {exception.Message}", exception);
            }
        }

        private static void ThrowIfUnsupportedTableStyleConditionalCellProperties(string styleId, LegacyDocElement cellProperties) {
            foreach (LegacyDocElement child in cellProperties.Elements()) {
                switch (child.WName) {
                    case "vAlign":
                        ReadSupportedTableCellVerticalAlignment(child);
                        break;
                    case "textDirection":
                        ReadSupportedTableCellTextDirection(child);
                        break;
                    case "tcFitText":
                        ReadSupportedTableCellFitText(child);
                        break;
                    case "noWrap":
                        ReadSupportedTableCellNoWrap(child);
                        break;
                    case "hideMark":
                        ReadSupportedTableCellHideMark(child);
                        break;
                    case "tcMar":
                        ReadSupportedTableCellMargins(child);
                        break;
                    case "shd":
                        ReadSupportedTableCellShading(child, "conditional table style shading");
                        break;
                    case "tcBorders":
                        ThrowIfUnsupportedTableStyleConditionalCellBorders(styleId, child);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional cell formatting only with supported layout, borders, and shading. Unsupported conditional cell property: {child.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableStyleConditionalCellBorders(string styleId, LegacyDocElement borders) {
            foreach (LegacyDocElement child in borders.Elements()) {
                switch (child.WName) {
                    case "top":
                    case "left":
                    case "bottom":
                    case "right":
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional borders only on top, left, bottom, and right cell edges. Unsupported conditional border: {child.LocalName}.");
                }
            }

            ReadSupportedTableCellBorder(borders.Child("top"));
            ReadSupportedTableCellBorder(borders.Child("left"));
            ReadSupportedTableCellBorder(borders.Child("bottom"));
            ReadSupportedTableCellBorder(borders.Child("right"));
        }

        private static void ThrowIfUnsupportedTableStyleConditionalTableProperties(string styleId, LegacyDocElement tableProperties) {
            foreach (LegacyDocElement child in tableProperties.Elements()) {
                switch (child.WName) {
                    case "tblBorders":
                        ReadSupportedTableBorders(child);
                        break;
                    case "shd":
                        ReadSupportedTableCellShading(child, "conditional table style shading");
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional table formatting only with supported borders and shading. Unsupported conditional table property: {child.LocalName}.");
                }
            }
        }
    }
}

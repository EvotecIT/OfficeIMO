using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const ushort CustomParagraphStyleSti = 0x0FFF;
        private const ushort NoBaseStyleIndex = 0x0FFF;
        private static readonly IReadOnlyDictionary<string, ushort> EmptyStyleIndexes = new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase);

        private static LegacyDocWritableStyleSheet CreateWritableStyleSheet(MainDocumentPart mainPart, Body body) {
            string[] usedStyleIds = body.Descendants<ParagraphStyleId>()
                .Select(style => style.Val?.Value)
                .Where(styleId => !string.IsNullOrWhiteSpace(styleId))
                .Select(styleId => styleId!)
                .Where(styleId => !string.Equals(styleId, "Normal", StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            ushort[] usedBuiltInStyleIndexes = usedStyleIds
                .Select(styleId => TryMapBuiltInParagraphStyleIndex(styleId, out ushort styleIndex) ? (ushort?)styleIndex : null)
                .Where(styleIndex => styleIndex != null)
                .Select(styleIndex => styleIndex!.Value)
                .Distinct()
                .OrderBy(styleIndex => styleIndex)
                .ToArray();
            string[] customStyleIds = usedStyleIds
                .Where(styleId => !TryMapBuiltInParagraphStyleIndex(styleId, out _))
                .ToArray();

            if (usedBuiltInStyleIndexes.Length == 0 && customStyleIds.Length == 0) {
                return LegacyDocWritableStyleSheet.Empty;
            }

            Styles? styles = mainPart.StyleDefinitionsPart?.Styles;
            if (styles == null && customStyleIds.Length > 0) {
                throw new NotSupportedException("Native DOC saving cannot write custom paragraph styles because the document has no style definitions part.");
            }

            Dictionary<string, Style> paragraphStyles = (styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                .Where(style => style.Type?.Value == StyleValues.Paragraph)
                .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                .GroupBy(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            var orderedStyleIds = new List<string>();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var visiting = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string styleId in customStyleIds) {
                AddCustomStyleWithBase(styleId, paragraphStyles, orderedStyleIds, visited, visiting);
            }

            Dictionary<ushort, Style> builtInStyles = ReadUsedBuiltInStyles(usedBuiltInStyleIndexes, paragraphStyles);
            var styleIndexes = new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < orderedStyleIds.Count; index++) {
                styleIndexes[orderedStyleIds[index]] = checked((ushort)(10 + index));
            }

            IReadOnlyList<string> fontFamilies = ReadStyleFontFamilies(builtInStyles.Values, orderedStyleIds, paragraphStyles);
            var fontFamilyIndexes = fontFamilies
                .Select((fontFamily, index) => new { fontFamily, index })
                .ToDictionary(item => item.fontFamily, item => item.index, StringComparer.OrdinalIgnoreCase);

            byte[] bytes = CreateWritableStyleSheetBytes(builtInStyles, orderedStyleIds, paragraphStyles, styleIndexes, fontFamilies, fontFamilyIndexes);
            return new LegacyDocWritableStyleSheet(bytes, styleIndexes, fontFamilies);
        }

        private static void AddCustomStyleWithBase(
            string styleId,
            IReadOnlyDictionary<string, Style> paragraphStyles,
            List<string> orderedStyleIds,
            HashSet<string> visited,
            HashSet<string> visiting) {
            if (string.Equals(styleId, "Normal", StringComparison.OrdinalIgnoreCase)
                || TryMapBuiltInParagraphStyleIndex(styleId, out _)) {
                return;
            }

            if (visited.Contains(styleId)) {
                return;
            }

            if (!visiting.Add(styleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write custom paragraph style '{styleId}' because its basedOn chain contains a cycle.");
            }

            if (!paragraphStyles.TryGetValue(styleId, out Style? style)) {
                throw new NotSupportedException($"Native DOC saving cannot write custom paragraph style '{styleId}' because the style definition is missing.");
            }

            string? basedOnStyleId = style.GetFirstChild<BasedOn>()?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(basedOnStyleId)) {
                AddCustomStyleWithBase(basedOnStyleId!, paragraphStyles, orderedStyleIds, visited, visiting);
            }

            visiting.Remove(styleId);
            visited.Add(styleId);
            orderedStyleIds.Add(styleId);
        }

        private static Dictionary<ushort, Style> ReadUsedBuiltInStyles(ushort[] usedBuiltInStyleIndexes, IReadOnlyDictionary<string, Style> paragraphStyles) {
            var builtInStyles = new Dictionary<ushort, Style>();
            foreach (ushort styleIndex in usedBuiltInStyleIndexes) {
                string styleId = GetBuiltInParagraphStyleId(styleIndex);
                if (!paragraphStyles.TryGetValue(styleId, out Style? style)) {
                    continue;
                }

                string? baseStyleId = style.GetFirstChild<BasedOn>()?.Val?.Value;
                if (!string.IsNullOrWhiteSpace(baseStyleId)
                    && !string.Equals(baseStyleId, "Normal", StringComparison.OrdinalIgnoreCase)
                    && !TryMapBuiltInParagraphStyleIndex(baseStyleId!, out _)) {
                    throw new NotSupportedException($"Native DOC saving cannot write built-in paragraph style '{styleId}' because basedOn style '{baseStyleId}' is not supported.");
                }

                _ = ReadSupportedBuiltInStyleParagraphFormatting(styleIndex, style.StyleParagraphProperties);
                _ = ReadSupportedRunFormatting(style.StyleRunProperties);
                builtInStyles[styleIndex] = style;
            }

            return builtInStyles;
        }

        private static IReadOnlyList<string> ReadStyleFontFamilies(IEnumerable<Style> builtInStyles, IReadOnlyList<string> styleIds, IReadOnlyDictionary<string, Style> paragraphStyles) {
            return builtInStyles
                .Concat(styleIds.Select(styleId => paragraphStyles[styleId]))
                .Select(style => style.StyleRunProperties)
                .Select(ReadSupportedRunFormatting)
                .Select(formatting => formatting.FontFamily)
                .Where(fontFamily => !string.IsNullOrWhiteSpace(fontFamily))
                .Select(fontFamily => fontFamily!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private static byte[] CreateWritableStyleSheetBytes(
            IReadOnlyDictionary<ushort, Style> builtInStyles,
            IReadOnlyList<string> customStyleIds,
            IReadOnlyDictionary<string, Style> paragraphStyles,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            IReadOnlyList<string> fontFamilies,
            IReadOnlyDictionary<string, int> fontFamilyIndexes) {
            var styleRecords = new List<byte[]>(10 + customStyleIds.Count) {
                CreateBuiltInParagraphStyleRecord(0, builtInStyles, fontFamilyIndexes)
            };

            for (ushort index = 1; index <= 9; index++) {
                styleRecords.Add(CreateBuiltInParagraphStyleRecord(index, builtInStyles, fontFamilyIndexes));
            }

            foreach (string styleId in customStyleIds) {
                Style style = paragraphStyles[styleId];
                ushort styleIndex = styleIndexes[styleId];
                ushort baseStyleIndex = ResolveBaseStyleIndex(style, styleIndexes);
                string styleName = ReadStyleName(style, styleId);
                LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedStyleParagraphFormatting(style.StyleParagraphProperties);
                LegacyDocWritableFormatting characterFormatting = ReadSupportedRunFormatting(style.StyleRunProperties);
                byte[] paragraphUpx = LegacyDocParagraphFormattingWriter.CreateStyleParagraphUpx(paragraphFormatting);
                byte[] characterUpx = CreateStyleCharacterUpx(characterFormatting, fontFamilyIndexes);
                styleRecords.Add(CreateParagraphStyleRecord(CustomParagraphStyleSti, baseStyleIndex, styleName, paragraphUpx, characterUpx));

                if (styleRecords.Count - 1 != styleIndex) {
                    throw new InvalidOperationException("The generated DOC stylesheet index map is inconsistent.");
                }
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, 4);
            WriteUInt16(stream, checked((ushort)styleRecords.Count));
            WriteUInt16(stream, 8);

            foreach (byte[] styleRecord in styleRecords) {
                WriteUInt16(stream, checked((ushort)styleRecord.Length));
                stream.Write(styleRecord, 0, styleRecord.Length);
                if ((stream.Position & 1) != 0) {
                    stream.WriteByte(0);
                }
            }

            return stream.ToArray();
        }

        private static byte[] CreateBuiltInParagraphStyleRecord(ushort styleIndex, IReadOnlyDictionary<ushort, Style> builtInStyles, IReadOnlyDictionary<string, int> fontFamilyIndexes) {
            if (!builtInStyles.TryGetValue(styleIndex, out Style? style)) {
                ushort baseStyleIndex = styleIndex == 0 ? NoBaseStyleIndex : (ushort)0;
                return CreateParagraphStyleRecord(styleIndex, baseStyleIndex, GetBuiltInParagraphStyleName(styleIndex), Array.Empty<byte>(), Array.Empty<byte>());
            }

            ushort basedOnStyleIndex = ResolveBuiltInBaseStyleIndex(style);
            string styleName = ReadStyleName(style, GetBuiltInParagraphStyleName(styleIndex));
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedBuiltInStyleParagraphFormatting(styleIndex, style.StyleParagraphProperties);
            LegacyDocWritableFormatting characterFormatting = ReadSupportedRunFormatting(style.StyleRunProperties);
            byte[] paragraphUpx = LegacyDocParagraphFormattingWriter.CreateStyleParagraphUpx(paragraphFormatting);
            byte[] characterUpx = CreateStyleCharacterUpx(characterFormatting, fontFamilyIndexes);
            return CreateParagraphStyleRecord(styleIndex, basedOnStyleIndex, styleName, paragraphUpx, characterUpx);
        }

        private static ushort ResolveBuiltInBaseStyleIndex(Style style) {
            string? baseStyleId = style.GetFirstChild<BasedOn>()?.Val?.Value;
            if (string.IsNullOrWhiteSpace(baseStyleId)) {
                return string.Equals(style.StyleId?.Value, "Normal", StringComparison.OrdinalIgnoreCase) ? NoBaseStyleIndex : (ushort)0;
            }

            if (string.Equals(baseStyleId, "Normal", StringComparison.OrdinalIgnoreCase)) {
                return 0;
            }

            if (TryMapBuiltInParagraphStyleIndex(baseStyleId!, out ushort builtInStyleIndex)) {
                return builtInStyleIndex;
            }

            throw new NotSupportedException($"Native DOC saving cannot write built-in paragraph style '{style.StyleId?.Value}' because basedOn style '{baseStyleId}' is not supported.");
        }

        private static ushort ResolveBaseStyleIndex(Style style, IReadOnlyDictionary<string, ushort> styleIndexes) {
            string? baseStyleId = style.GetFirstChild<BasedOn>()?.Val?.Value;
            if (string.IsNullOrWhiteSpace(baseStyleId)) {
                return 0;
            }

            if (string.Equals(baseStyleId, "Normal", StringComparison.OrdinalIgnoreCase)) {
                return 0;
            }

            if (TryMapBuiltInParagraphStyleIndex(baseStyleId!, out ushort builtInStyleIndex)) {
                return builtInStyleIndex;
            }

            if (styleIndexes.TryGetValue(baseStyleId!, out ushort customStyleIndex)) {
                return customStyleIndex;
            }

            throw new NotSupportedException($"Native DOC saving cannot write custom paragraph style '{style.StyleId?.Value}' because basedOn style '{baseStyleId}' is not supported.");
        }

        private static string ReadStyleName(Style style, string styleId) {
            string? name = style.StyleName?.Val?.Value;
            return string.IsNullOrWhiteSpace(name) ? styleId : name!;
        }

        private static string GetBuiltInParagraphStyleId(ushort styleIndex) {
            if (styleIndex == 0) {
                return "Normal";
            }

            return "Heading" + styleIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string GetBuiltInParagraphStyleName(ushort styleIndex) {
            if (styleIndex == 0) {
                return "Normal";
            }

            return "heading " + styleIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static byte[] CreateParagraphStyleRecord(ushort sti, ushort baseStyleIndex, string name, byte[] paragraphUpx, byte[] characterUpx) {
            var upxs = characterUpx.Length == 0
                ? paragraphUpx.Length == 0
                    ? Array.Empty<byte[]>()
                    : new[] { paragraphUpx }
                : new[] { paragraphUpx, characterUpx };

            using var stream = new MemoryStream();
            WriteUInt16(stream, sti);
            WriteUInt16(stream, checked((ushort)((baseStyleIndex << 4) | 1)));
            WriteUInt16(stream, checked((ushort)upxs.Length));
            WriteUInt16(stream, 0);
            WriteXstz(stream, name);

            foreach (byte[] upx in upxs) {
                WriteUInt16(stream, checked((ushort)upx.Length));
                stream.Write(upx, 0, upx.Length);
                if ((stream.Position & 1) != 0) {
                    stream.WriteByte(0);
                }
            }

            return stream.ToArray();
        }

        private static void WriteXstz(Stream stream, string value) {
            if (value.Length > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write a custom paragraph style whose name is longer than the DOC stylesheet limit.");
            }

            WriteUInt16(stream, checked((ushort)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
            WriteUInt16(stream, 0);
        }

        private readonly struct LegacyDocWritableStyleSheet {
            internal static readonly LegacyDocWritableStyleSheet Empty = new LegacyDocWritableStyleSheet(
                Array.Empty<byte>(),
                EmptyStyleIndexes,
                Array.Empty<string>());

            internal LegacyDocWritableStyleSheet(byte[] bytes, IReadOnlyDictionary<string, ushort> styleIndexes, IReadOnlyList<string> fontFamilies) {
                Bytes = bytes;
                StyleIndexes = styleIndexes;
                FontFamilies = fontFamilies;
            }

            internal byte[] Bytes { get; }

            internal IReadOnlyDictionary<string, ushort> StyleIndexes { get; }

            internal IReadOnlyList<string> FontFamilies { get; }
        }
    }
}

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
            string[] customStyleIds = body.Descendants<ParagraphStyleId>()
                .Select(style => style.Val?.Value)
                .Where(styleId => !string.IsNullOrWhiteSpace(styleId))
                .Select(styleId => styleId!)
                .Where(styleId => !string.Equals(styleId, "Normal", StringComparison.OrdinalIgnoreCase))
                .Where(styleId => !TryMapBuiltInParagraphStyleIndex(styleId, out _))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            if (customStyleIds.Length == 0) {
                return LegacyDocWritableStyleSheet.Empty;
            }

            Styles? styles = mainPart.StyleDefinitionsPart?.Styles;
            if (styles == null) {
                throw new NotSupportedException("Native DOC saving cannot write custom paragraph styles because the document has no style definitions part.");
            }

            Dictionary<string, Style> paragraphStyles = styles.Elements<Style>()
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

            var styleIndexes = new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < orderedStyleIds.Count; index++) {
                styleIndexes[orderedStyleIds[index]] = checked((ushort)(10 + index));
            }

            IReadOnlyList<string> fontFamilies = ReadStyleFontFamilies(orderedStyleIds, paragraphStyles);
            var fontFamilyIndexes = fontFamilies
                .Select((fontFamily, index) => new { fontFamily, index })
                .ToDictionary(item => item.fontFamily, item => item.index, StringComparer.OrdinalIgnoreCase);

            byte[] bytes = CreateWritableStyleSheetBytes(orderedStyleIds, paragraphStyles, styleIndexes, fontFamilies, fontFamilyIndexes);
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

        private static IReadOnlyList<string> ReadStyleFontFamilies(IReadOnlyList<string> styleIds, IReadOnlyDictionary<string, Style> paragraphStyles) {
            return styleIds
                .Select(styleId => paragraphStyles[styleId].StyleRunProperties)
                .Select(ReadSupportedRunFormatting)
                .Select(formatting => formatting.FontFamily)
                .Where(fontFamily => !string.IsNullOrWhiteSpace(fontFamily))
                .Select(fontFamily => fontFamily!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private static byte[] CreateWritableStyleSheetBytes(
            IReadOnlyList<string> customStyleIds,
            IReadOnlyDictionary<string, Style> paragraphStyles,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            IReadOnlyList<string> fontFamilies,
            IReadOnlyDictionary<string, int> fontFamilyIndexes) {
            var styleRecords = new List<byte[]>(10 + customStyleIds.Count) {
                CreateParagraphStyleRecord(0, NoBaseStyleIndex, "Normal", Array.Empty<byte>(), Array.Empty<byte>())
            };

            for (ushort index = 1; index <= 9; index++) {
                styleRecords.Add(CreateParagraphStyleRecord(index, 0, "heading " + index.ToString(System.Globalization.CultureInfo.InvariantCulture), Array.Empty<byte>(), Array.Empty<byte>()));
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

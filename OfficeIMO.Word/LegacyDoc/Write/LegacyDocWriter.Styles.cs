using OfficeIMO.Core.Internal.OpenXml;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const ushort CustomParagraphStyleSti = 0x0FFE;
        private const ushort NoBaseStyleIndex = 0x0FFF;
        private static readonly IReadOnlyDictionary<string, ushort> EmptyStyleIndexes = new Dictionary<string, ushort>(StringComparer.Ordinal);

        private static LegacyDocWritableStyleSheet CreateWritableStyleSheet(LegacyDocSourcePackage source, LegacyDocElement body) {
            // Header/footer, note and comment stories share the document's style sheet.
            // Discover their referenced IDs, including custom IDs produced by native import.
            LegacyDocElement[] sections = body.Descendants("sectPr").Where(IsParagraphOrBodySectionProperties).ToArray();
            foreach (LegacyDocElement section in sections) {
                LegacyDocElement[] headers = section.Elements("headerReference").ToArray();
                LegacyDocElement[] footers = section.Elements("footerReference").ToArray();
                foreach (string type in new[] { "default", "first", "even" }) {
                    GetSingleReference(headers, type, "header");
                    GetSingleReference(footers, type, "footer");
                }
            }
            IEnumerable<LegacyDocElement> storyStyleIds = sections.SelectMany(section => section.Elements("headerReference"))
                .Select(reference => reference.Attr(OfficeXmlNamespaces.Relationships, "id")).Distinct(StringComparer.Ordinal)
                .Select(id => GetReferencedHeaderOrFooterPart(source, id, "header", header: true)).Distinct()
                .SelectMany(part => ParagraphStyleIds(part.Root))
                .Concat(sections.SelectMany(section => section.Elements("footerReference"))
                    .Select(reference => reference.Attr(OfficeXmlNamespaces.Relationships, "id")).Distinct(StringComparer.Ordinal)
                    .Select(id => GetReferencedHeaderOrFooterPart(source, id, "footer", header: false)).Distinct()
                    .SelectMany(part => ParagraphStyleIds(part.Root)))
                .Concat(source.Footnotes.Elements("footnote")
                    .Where(IsUserFootnote)
                    .SelectMany(ParagraphStyleIds))
                .Concat(source.Endnotes.Elements("endnote")
                    .Where(IsUserEndnote)
                    .SelectMany(ParagraphStyleIds))
                .Concat(ParagraphStyleIds(source.Comments));
            string[] usedStyleIds = ParagraphStyleIds(body).Concat(storyStyleIds)
                .Select(style => style.Val)
                .Where(styleId => !string.IsNullOrWhiteSpace(styleId))
                .Select(styleId => styleId!)
                .Where(styleId => !string.Equals(styleId, "Normal", StringComparison.Ordinal))
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            var builtInStyleIndexes = new SortedSet<ushort>(usedStyleIds
                .Select(styleId => TryMapBuiltInParagraphStyleIndex(styleId, out ushort styleIndex) ? (ushort?)styleIndex : null)
                .Where(styleIndex => styleIndex != null)
                .Select(styleIndex => styleIndex!.Value));
            string[] customStyleIds = usedStyleIds
                .Where(styleId => !TryMapBuiltInParagraphStyleIndex(styleId, out _))
                .ToArray();

            builtInStyleIndexes.Add(0);

            LegacyDocStyleCatalog styles = source.Styles;
            if (!styles.Exists && customStyleIds.Length > 0) {
                throw new NotSupportedException("Native DOC saving cannot write custom paragraph styles because the document has no style definitions part.");
            }

            // The materializing steps below replace entries with detached copies carrying the effective definitions.
            Dictionary<string, LegacyDocElement> paragraphStyles = styles.Definitions
                .Where(LegacyDocStyleCatalog.IsParagraphStyle)
                .Where(style => !string.IsNullOrWhiteSpace(LegacyDocStyleCatalog.StyleId(style)))
                .GroupBy(style => LegacyDocStyleCatalog.StyleId(style)!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            LegacyDocElement documentDefaults = styles.DocumentDefaults;
            paragraphStyles["Normal"] = CreateDefaultParagraphStyle(source, paragraphStyles, documentDefaults);
            MaterializeParagraphStyleThemeFonts(paragraphStyles, source);
            MaterializeDocumentDefaultSpacing(paragraphStyles, documentDefaults);
            MaterializeDocumentDefaultCharacterScale(paragraphStyles, documentDefaults);

            var orderedStyleIds = new List<string>();
            var visited = new HashSet<string>(StringComparer.Ordinal);
            var visiting = new HashSet<string>(StringComparer.Ordinal);
            foreach (string styleId in customStyleIds) {
                AddCustomStyleWithBase(styleId, paragraphStyles, orderedStyleIds, builtInStyleIndexes, visited, visiting);
            }

            AddBuiltInBaseStyles(builtInStyleIndexes, paragraphStyles);
            Dictionary<ushort, LegacyDocElement> builtInStyles = ReadUsedBuiltInStyles(builtInStyleIndexes, paragraphStyles);
            var styleIndexes = new Dictionary<string, ushort>(StringComparer.Ordinal);
            for (int index = 0; index < orderedStyleIds.Count; index++) {
                styleIndexes[orderedStyleIds[index]] = checked((ushort)(15 + index));
            }

            IReadOnlyList<string> fontFamilies = ReadStyleFontFamilies(builtInStyles.Values, orderedStyleIds, paragraphStyles);
            var fontFamilyIndexes = fontFamilies
                .Select((fontFamily, index) => new { fontFamily, index })
                .ToDictionary(item => item.fontFamily, item => item.index, StringComparer.OrdinalIgnoreCase);

            byte[] bytes = CreateWritableStyleSheetBytes(builtInStyles, orderedStyleIds, paragraphStyles, styleIndexes, fontFamilies, fontFamilyIndexes);
            return new LegacyDocWritableStyleSheet(bytes, styleIndexes, fontFamilies);
        }

        // The SDK read w:sectPr as section properties only on the body and in a paragraph's own properties; elsewhere
        // (inside a tracked section change) it was a different element.
        private static bool IsParagraphOrBodySectionProperties(LegacyDocElement sectionProperties) {
            LegacyDocElement parent = sectionProperties.Parent;
            return parent.Is("body") || (parent.Is("pPr") && parent.Parent.Is("p"));
        }

        /// <summary>The paragraph style references (<c>w:pStyle</c> in paragraph properties) below an element.</summary>
        private static IEnumerable<LegacyDocElement> ParagraphStyleIds(LegacyDocElement root) =>
            root.Descendants("pStyle").Where(styleId => styleId.Parent.Is("pPr"));

        private static void AddCustomStyleWithBase(
            string styleId,
            IReadOnlyDictionary<string, LegacyDocElement> paragraphStyles,
            List<string> orderedStyleIds,
            ISet<ushort> builtInStyleIndexes,
            HashSet<string> visited,
            HashSet<string> visiting) {
            if (string.Equals(styleId, "Normal", StringComparison.Ordinal)) {
                return;
            }

            if (TryMapBuiltInParagraphStyleIndex(styleId, out ushort builtInStyleIndex)) {
                builtInStyleIndexes.Add(builtInStyleIndex);
                return;
            }

            if (visited.Contains(styleId)) {
                return;
            }

            if (!visiting.Add(styleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write custom paragraph style '{styleId}' because its basedOn chain contains a cycle.");
            }

            if (!paragraphStyles.TryGetValue(styleId, out LegacyDocElement style)) {
                throw new NotSupportedException($"Native DOC saving cannot write custom paragraph style '{styleId}' because the style definition is missing.");
            }

            string? basedOnStyleId = style.Child("basedOn").Val;
            if (!string.IsNullOrWhiteSpace(basedOnStyleId)) {
                AddCustomStyleWithBase(basedOnStyleId!, paragraphStyles, orderedStyleIds, builtInStyleIndexes, visited, visiting);
            }

            string? nextStyleId = style.Child("next").Val;
            if (!string.IsNullOrWhiteSpace(nextStyleId) && !string.Equals(nextStyleId, styleId, StringComparison.Ordinal)) {
                AddCustomStyleWithBase(nextStyleId!, paragraphStyles, orderedStyleIds, builtInStyleIndexes, visited, visiting);
            }

            visiting.Remove(styleId);
            visited.Add(styleId);
            orderedStyleIds.Add(styleId);
        }

        private static void AddBuiltInBaseStyles(ISet<ushort> builtInStyleIndexes, IReadOnlyDictionary<string, LegacyDocElement> paragraphStyles) {
            var pending = new Queue<ushort>(builtInStyleIndexes);
            while (pending.Count > 0) {
                ushort styleIndex = pending.Dequeue();
                string styleId = GetBuiltInParagraphStyleId(styleIndex);
                if (!paragraphStyles.TryGetValue(styleId, out LegacyDocElement style)) {
                    continue;
                }

                string? baseStyleId = style.Child("basedOn").Val;
                if (!string.IsNullOrWhiteSpace(baseStyleId)
                    && !string.Equals(baseStyleId, "Normal", StringComparison.Ordinal)
                    && TryMapBuiltInParagraphStyleIndex(baseStyleId!, out ushort baseStyleIndex)
                    && builtInStyleIndexes.Add(baseStyleIndex)) {
                    pending.Enqueue(baseStyleIndex);
                }

                string? nextStyleId = style.Child("next").Val;
                if (!string.IsNullOrWhiteSpace(nextStyleId)
                    && !string.Equals(nextStyleId, "Normal", StringComparison.Ordinal)
                    && TryMapBuiltInParagraphStyleIndex(nextStyleId!, out ushort nextStyleIndex)
                    && builtInStyleIndexes.Add(nextStyleIndex)) {
                    pending.Enqueue(nextStyleIndex);
                }
            }
        }

        private static Dictionary<ushort, LegacyDocElement> ReadUsedBuiltInStyles(IEnumerable<ushort> usedBuiltInStyleIndexes, IReadOnlyDictionary<string, LegacyDocElement> paragraphStyles) {
            var builtInStyles = new Dictionary<ushort, LegacyDocElement>();
            foreach (ushort styleIndex in usedBuiltInStyleIndexes) {
                string styleId = GetBuiltInParagraphStyleId(styleIndex);
                if (!paragraphStyles.TryGetValue(styleId, out LegacyDocElement style)) {
                    continue;
                }

                string? baseStyleId = style.Child("basedOn").Val;
                if (!string.IsNullOrWhiteSpace(baseStyleId)
                    && !string.Equals(baseStyleId, "Normal", StringComparison.Ordinal)
                    && !TryMapBuiltInParagraphStyleIndex(baseStyleId!, out _)) {
                    throw new NotSupportedException($"Native DOC saving cannot write built-in paragraph style '{styleId}' because basedOn style '{baseStyleId}' is not supported.");
                }

                _ = ReadSupportedBuiltInStyleParagraphFormatting(styleIndex, style.Child("pPr"));
                _ = ReadSupportedStyleRunFormatting(style);
                builtInStyles[styleIndex] = style;
            }

            return builtInStyles;
        }

        private static LegacyDocWritableFormatting ReadSupportedStyleRunFormatting(LegacyDocElement style) {
            if (style.Child("rPr").Child("highlight").Exists) {
                throw new NotSupportedException($"Native DOC saving cannot write paragraph style '{style.Attr("styleId")}' because highlight is not permitted in DOC styles.");
            }

            return ReadSupportedRunFormatting(style.Child("rPr"));
        }

        private static IReadOnlyList<string> ReadStyleFontFamilies(IEnumerable<LegacyDocElement> builtInStyles, IReadOnlyList<string> styleIds, IReadOnlyDictionary<string, LegacyDocElement> paragraphStyles) {
            return builtInStyles
                .Concat(styleIds.Select(styleId => paragraphStyles[styleId]))
                .Select(ReadSupportedStyleRunFormatting)
                .Select(formatting => formatting.FontFamily)
                .Where(fontFamily => !string.IsNullOrWhiteSpace(fontFamily))
                .Select(fontFamily => fontFamily!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private static byte[] CreateWritableStyleSheetBytes(
            IReadOnlyDictionary<ushort, LegacyDocElement> builtInStyles,
            IReadOnlyList<string> customStyleIds,
            IReadOnlyDictionary<string, LegacyDocElement> paragraphStyles,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            IReadOnlyList<string> fontFamilies,
            IReadOnlyDictionary<string, int> fontFamilyIndexes) {
            var styleRecords = new List<byte[]>(15 + customStyleIds.Count) {
                CreateBuiltInParagraphStyleRecord(0, builtInStyles, fontFamilyIndexes)
            };

            for (ushort index = 1; index <= 9; index++) {
                styleRecords.Add(CreateBuiltInParagraphStyleRecord(index, builtInStyles, fontFamilyIndexes));
            }
            for (int index = 10; index < 15; index++) styleRecords.Add(Array.Empty<byte>());

            foreach (string styleId in customStyleIds) {
                LegacyDocElement style = paragraphStyles[styleId];
                ushort styleIndex = styleIndexes[styleId];
                ushort baseStyleIndex = ResolveBaseStyleIndex(style, styleIndexes);
                ushort nextStyleIndex = ResolveNextStyleIndex(style, styleIndex, styleIndexes);
                string styleName = ReadStyleName(style, styleId);
                LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedCustomParagraphStyleParagraphFormatting(style.Child("pPr"));
                LegacyDocWritableFormatting characterFormatting = ReadSupportedStyleRunFormatting(style);
                byte[] paragraphUpx = LegacyDocParagraphFormattingWriter.CreateStyleParagraphUpx(paragraphFormatting);
                byte[] characterUpx = CreateStyleCharacterUpx(characterFormatting, fontFamilyIndexes);
                styleRecords.Add(CreateParagraphStyleRecord(CustomParagraphStyleSti, baseStyleIndex, nextStyleIndex, styleName, paragraphUpx, characterUpx, styleIndex));

                if (styleRecords.Count - 1 != styleIndex) {
                    throw new InvalidOperationException("The generated DOC stylesheet index map is inconsistent.");
                }
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, 18);
            WriteUInt16(stream, checked((ushort)styleRecords.Count));
            WriteUInt16(stream, 10);
            WriteUInt16(stream, 1); // Built-in style names are written.
            WriteUInt16(stream, 10);
            WriteUInt16(stream, 15); // Reserved fixed-index styles.
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 0); // Default ASCII, East Asian and other font indexes.
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 0);

            foreach (byte[] styleRecord in styleRecords) {
                WriteUInt16(stream, checked((ushort)styleRecord.Length));
                stream.Write(styleRecord, 0, styleRecord.Length);
                if ((stream.Position & 1) != 0) {
                    stream.WriteByte(0);
                }
            }

            return stream.ToArray();
        }

        private static byte[] CreateBuiltInParagraphStyleRecord(ushort styleIndex, IReadOnlyDictionary<ushort, LegacyDocElement> builtInStyles, IReadOnlyDictionary<string, int> fontFamilyIndexes) {
            if (!builtInStyles.TryGetValue(styleIndex, out LegacyDocElement style)) {
                ushort baseStyleIndex = styleIndex == 0 ? NoBaseStyleIndex : (ushort)0;
                return CreateParagraphStyleRecord(styleIndex, baseStyleIndex, 0, GetBuiltInParagraphStyleName(styleIndex), Array.Empty<byte>(), Array.Empty<byte>());
            }

            ushort basedOnStyleIndex = ResolveBuiltInBaseStyleIndex(style);
            ushort nextStyleIndex = ResolveNextStyleIndex(style, styleIndex, EmptyStyleIndexes);
            string styleName = ReadStyleName(style, GetBuiltInParagraphStyleName(styleIndex));
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedBuiltInStyleParagraphFormatting(styleIndex, style.Child("pPr"));
            LegacyDocWritableFormatting characterFormatting = ReadSupportedStyleRunFormatting(style);
            byte[] paragraphUpx = LegacyDocParagraphFormattingWriter.CreateStyleParagraphUpx(paragraphFormatting);
            byte[] characterUpx = CreateStyleCharacterUpx(characterFormatting, fontFamilyIndexes);
            return CreateParagraphStyleRecord(styleIndex, basedOnStyleIndex, nextStyleIndex, styleName, paragraphUpx, characterUpx);
        }

        private static ushort ResolveBuiltInBaseStyleIndex(LegacyDocElement style) {
            string? baseStyleId = style.Child("basedOn").Val;
            if (string.IsNullOrWhiteSpace(baseStyleId)) {
                return string.Equals(style.Attr("styleId"), "Normal", StringComparison.Ordinal) ? NoBaseStyleIndex : (ushort)0;
            }

            if (string.Equals(baseStyleId, "Normal", StringComparison.Ordinal)) {
                return 0;
            }

            if (TryMapBuiltInParagraphStyleIndex(baseStyleId!, out ushort builtInStyleIndex)) {
                return builtInStyleIndex;
            }

            throw new NotSupportedException($"Native DOC saving cannot write built-in paragraph style '{style.Attr("styleId")}' because basedOn style '{baseStyleId}' is not supported.");
        }

        private static ushort ResolveBaseStyleIndex(LegacyDocElement style, IReadOnlyDictionary<string, ushort> styleIndexes) {
            string? baseStyleId = style.Child("basedOn").Val;
            if (string.IsNullOrWhiteSpace(baseStyleId)) {
                return 0;
            }

            if (string.Equals(baseStyleId, "Normal", StringComparison.Ordinal)) {
                return 0;
            }

            if (TryMapBuiltInParagraphStyleIndex(baseStyleId!, out ushort builtInStyleIndex)) {
                return builtInStyleIndex;
            }

            if (styleIndexes.TryGetValue(baseStyleId!, out ushort customStyleIndex)) {
                return customStyleIndex;
            }

            throw new NotSupportedException($"Native DOC saving cannot write custom paragraph style '{style.Attr("styleId")}' because basedOn style '{baseStyleId}' is not supported.");
        }

        private static ushort ResolveNextStyleIndex(LegacyDocElement style, ushort currentStyleIndex, IReadOnlyDictionary<string, ushort> styleIndexes) {
            string? nextStyleId = style.Child("next").Val;
            if (string.IsNullOrWhiteSpace(nextStyleId)) {
                return 0;
            }

            if (string.Equals(nextStyleId, style.Attr("styleId"), StringComparison.Ordinal)) {
                return currentStyleIndex;
            }

            if (string.Equals(nextStyleId, "Normal", StringComparison.Ordinal)) {
                return 0;
            }

            if (TryMapBuiltInParagraphStyleIndex(nextStyleId!, out ushort builtInStyleIndex)) {
                return builtInStyleIndex;
            }

            if (styleIndexes.TryGetValue(nextStyleId!, out ushort customStyleIndex)) {
                return customStyleIndex;
            }

            throw new NotSupportedException($"Native DOC saving cannot write paragraph style '{style.Attr("styleId")}' because next paragraph style '{nextStyleId}' is not supported.");
        }

        private static string ReadStyleName(LegacyDocElement style, string styleId) {
            string? name = style.Child("name").Val;
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

        private static byte[] CreateParagraphStyleRecord(ushort sti, ushort baseStyleIndex, ushort nextStyleIndex, string name, byte[] paragraphUpx, byte[] characterUpx, ushort? styleIndex = null) {
            if (paragraphUpx.Length == 0) paragraphUpx = new byte[2];
            ushort index = styleIndex ?? sti;
            paragraphUpx[0] = (byte)index;
            paragraphUpx[1] = (byte)(index >> 8);
            var upxs = new[] { paragraphUpx, characterUpx };

            using var stream = new MemoryStream();
            WriteUInt16(stream, sti);
            WriteUInt16(stream, checked((ushort)((baseStyleIndex << 4) | 1)));
            WriteUInt16(stream, checked((ushort)((nextStyleIndex << 4) | upxs.Length)));
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 0); // StdfBase flags, following bchUpe.
            WriteXstz(stream, name);

            foreach (byte[] upx in upxs) {
                WriteUInt16(stream, checked((ushort)upx.Length));
                stream.Write(upx, 0, upx.Length);
                if ((stream.Position & 1) != 0) {
                    stream.WriteByte(0);
                }
            }

            byte[] record = stream.ToArray();
            record[6] = (byte)record.Length;
            record[7] = (byte)(record.Length >> 8);
            return record;
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

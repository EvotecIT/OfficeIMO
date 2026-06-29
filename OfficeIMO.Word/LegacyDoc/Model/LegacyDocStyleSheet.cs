using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocStyleSheet {
        private readonly IReadOnlyDictionary<ushort, LegacyDocParagraphStyle> _paragraphStyles;

        private LegacyDocStyleSheet(IReadOnlyDictionary<ushort, LegacyDocParagraphStyle> paragraphStyles) {
            _paragraphStyles = paragraphStyles;
        }

        internal static LegacyDocStyleSheet Empty { get; } = new LegacyDocStyleSheet(new Dictionary<ushort, LegacyDocParagraphStyle>());

        internal IEnumerable<LegacyDocParagraphStyle> ParagraphStyles => _paragraphStyles.Values;

        internal bool TryGetParagraphStyle(ushort styleIndex, out LegacyDocParagraphStyle style) {
            return _paragraphStyles.TryGetValue(styleIndex, out style!);
        }

        internal static LegacyDocStyleSheet Read(byte[] tableStream, LegacyDocFib fib, out string? warning) {
            warning = null;
            if (fib.LcbStshf == 0) {
                return Empty;
            }

            if (fib.LcbStshf < 4
                || fib.FcStshf < 0
                || fib.FcStshf + fib.LcbStshf > tableStream.Length) {
                warning = "The FIB points outside the selected table stream for the stylesheet.";
                return Empty;
            }

            int offset = fib.FcStshf;
            int end = offset + fib.LcbStshf;
            int cbStshi = LegacyDocFib.ReadUInt16(tableStream, offset);
            int stshifOffset = offset + 2;
            if (cbStshi < 4 || stshifOffset + cbStshi > end) {
                warning = "The DOC stylesheet header points outside the stylesheet.";
                return Empty;
            }

            int cstd = LegacyDocFib.ReadUInt16(tableStream, stshifOffset);
            int cbStdBaseInFile = LegacyDocFib.ReadUInt16(tableStream, stshifOffset + 2);
            if (cstd < 0 || cbStdBaseInFile < 8) {
                warning = "The DOC stylesheet header contains an unsupported standard style base size.";
                return Empty;
            }

            offset = stshifOffset + cbStshi;
            var styles = new Dictionary<ushort, LegacyDocParagraphStyle>();
            var usedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (ushort styleIndex = 0; styleIndex < cstd && offset + 2 <= end; styleIndex++) {
                int cbStd = LegacyDocFib.ReadUInt16(tableStream, offset);
                offset += 2;
                if (cbStd == 0) {
                    continue;
                }

                if (offset + cbStd > end) {
                    warning = "A DOC stylesheet style record points outside the stylesheet.";
                    break;
                }

                if (TryReadParagraphStyle(tableStream, offset, cbStd, cbStdBaseInFile, styleIndex, usedStyleIds, out LegacyDocParagraphStyle? style)) {
                    styles[styleIndex] = style!;
                }

                offset += cbStd;
                if ((offset & 1) != 0) {
                    offset++;
                }
            }

            return styles.Count == 0 ? Empty : new LegacyDocStyleSheet(styles);
        }

        private static bool TryReadParagraphStyle(
            byte[] bytes,
            int offset,
            int count,
            int cbStdBaseInFile,
            ushort styleIndex,
            HashSet<string> usedStyleIds,
            out LegacyDocParagraphStyle? style) {
            style = null;
            if (count < cbStdBaseInFile) {
                return false;
            }

            ushort first = LegacyDocFib.ReadUInt16(bytes, offset);
            ushort second = LegacyDocFib.ReadUInt16(bytes, offset + 2);
            ushort sti = (ushort)(first & 0x0FFF);
            ushort stk = (ushort)(second & 0x000F);
            if (stk != 1) {
                return false;
            }

            int nameOffset = offset + cbStdBaseInFile;
            int end = offset + count;
            string? name = ReadXstz(bytes, nameOffset, end);
            if (string.IsNullOrWhiteSpace(name)) {
                return false;
            }

            if (TryMapBuiltInParagraphStyle(sti, name!, out WordParagraphStyles builtInStyle)) {
                style = LegacyDocParagraphStyle.ForBuiltIn(styleIndex, name!, builtInStyle);
                return true;
            }

            string styleId = CreateCustomStyleId(name!, styleIndex, usedStyleIds);
            style = LegacyDocParagraphStyle.ForCustom(styleIndex, name!, styleId);
            return true;
        }

        private static string? ReadXstz(byte[] bytes, int offset, int end) {
            if (offset + 2 <= end) {
                int charCount = LegacyDocFib.ReadUInt16(bytes, offset);
                int byteCount = charCount * 2;
                int textOffset = offset + 2;
                int terminatorOffset = textOffset + byteCount;
                if (charCount > 0
                    && terminatorOffset + 2 <= end
                    && bytes[terminatorOffset] == 0
                    && bytes[terminatorOffset + 1] == 0) {
                    return Encoding.Unicode.GetString(bytes, textOffset, byteCount);
                }
            }

            if (offset + 1 <= end) {
                int charCount = bytes[offset];
                int byteCount = charCount * 2;
                int textOffset = offset + 1;
                int terminatorOffset = textOffset + byteCount;
                if (charCount > 0
                    && terminatorOffset + 2 <= end
                    && bytes[terminatorOffset] == 0
                    && bytes[terminatorOffset + 1] == 0) {
                    return Encoding.Unicode.GetString(bytes, textOffset, byteCount);
                }
            }

            return null;
        }

        private static bool TryMapBuiltInParagraphStyle(ushort sti, string name, out WordParagraphStyles style) {
            if (TryMapBuiltInParagraphStyleIndex(sti, out style)) {
                return true;
            }

            string normalized = NormalizeStyleName(name);
            switch (normalized) {
                case "normal":
                    style = WordParagraphStyles.Normal;
                    return true;
                case "heading1":
                    style = WordParagraphStyles.Heading1;
                    return true;
                case "heading2":
                    style = WordParagraphStyles.Heading2;
                    return true;
                case "heading3":
                    style = WordParagraphStyles.Heading3;
                    return true;
                case "heading4":
                    style = WordParagraphStyles.Heading4;
                    return true;
                case "heading5":
                    style = WordParagraphStyles.Heading5;
                    return true;
                case "heading6":
                    style = WordParagraphStyles.Heading6;
                    return true;
                case "heading7":
                    style = WordParagraphStyles.Heading7;
                    return true;
                case "heading8":
                    style = WordParagraphStyles.Heading8;
                    return true;
                case "heading9":
                    style = WordParagraphStyles.Heading9;
                    return true;
                default:
                    style = default;
                    return false;
            }
        }

        private static bool TryMapBuiltInParagraphStyleIndex(ushort sti, out WordParagraphStyles style) {
            switch (sti) {
                case 0:
                    style = WordParagraphStyles.Normal;
                    return true;
                case 1:
                    style = WordParagraphStyles.Heading1;
                    return true;
                case 2:
                    style = WordParagraphStyles.Heading2;
                    return true;
                case 3:
                    style = WordParagraphStyles.Heading3;
                    return true;
                case 4:
                    style = WordParagraphStyles.Heading4;
                    return true;
                case 5:
                    style = WordParagraphStyles.Heading5;
                    return true;
                case 6:
                    style = WordParagraphStyles.Heading6;
                    return true;
                case 7:
                    style = WordParagraphStyles.Heading7;
                    return true;
                case 8:
                    style = WordParagraphStyles.Heading8;
                    return true;
                case 9:
                    style = WordParagraphStyles.Heading9;
                    return true;
                default:
                    style = default;
                    return false;
            }
        }

        private static string CreateCustomStyleId(string name, ushort styleIndex, HashSet<string> usedStyleIds) {
            string cleaned = new string(name.Where(char.IsLetterOrDigit).ToArray());
            if (string.IsNullOrEmpty(cleaned)) {
                cleaned = "Style";
            }

            string baseId = "LegacyDoc" + cleaned;
            string styleId = baseId;
            if (usedStyleIds.Add(styleId)) {
                return styleId;
            }

            styleId = baseId + styleIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
            usedStyleIds.Add(styleId);
            return styleId;
        }

        private static string NormalizeStyleName(string name) {
            return new string(name.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
        }
    }

    internal sealed class LegacyDocParagraphStyle {
        private LegacyDocParagraphStyle(ushort index, string name, string? styleId, WordParagraphStyles? builtInStyle) {
            Index = index;
            Name = name;
            StyleId = styleId;
            BuiltInStyle = builtInStyle;
        }

        internal ushort Index { get; }

        internal string Name { get; }

        internal string? StyleId { get; }

        internal WordParagraphStyles? BuiltInStyle { get; }

        internal static LegacyDocParagraphStyle ForBuiltIn(ushort index, string name, WordParagraphStyles style) {
            return new LegacyDocParagraphStyle(index, name, null, style);
        }

        internal static LegacyDocParagraphStyle ForCustom(ushort index, string name, string styleId) {
            return new LegacyDocParagraphStyle(index, name, styleId, null);
        }
    }
}

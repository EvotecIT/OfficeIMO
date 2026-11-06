namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed partial class LegacyDocStyleSheet {
        /// <summary>Reads ANSI Word 6/95 paragraph style records and projects their inherited formatting.</summary>
        internal static LegacyDocStyleSheet ReadWord69(ReadOnlySpan<byte> data, IReadOnlyList<string> fonts,
            LegacyWord69Properties properties) {
            if (data.IsEmpty) return Empty;
            if (data.Length < 6) throw new InvalidDataException("The Word 6/95 stylesheet is truncated.");

            int headerLength = LegacyWord69Reader.U16(data, 0);
            int count = LegacyWord69Reader.U16(data, 2);
            int baseLength = LegacyWord69Reader.U16(data, 4);
            int offset = 2 + headerLength;
            if (headerLength < 4 || offset > data.Length || baseLength < 8 || count > (data.Length - offset) / 2) {
                throw new InvalidDataException("The Word 6/95 stylesheet header is invalid.");
            }

            var styles = new Dictionary<ushort, LegacyDocParagraphStyle>();
            var usedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string? defaultFont = null;
            if (headerLength >= 14) {
                int defaultFontIndex = LegacyWord69Reader.U16(data, 14);
                if (defaultFontIndex < fonts.Count) {
                    defaultFont = fonts[defaultFontIndex];
                    properties.DefaultFontIndex = defaultFontIndex;
                }
            }

            for (ushort index = 0; index < count; index++) {
                if (offset > data.Length - 2) throw new InvalidDataException("The Word 6/95 style directory is truncated.");
                int length = LegacyWord69Reader.U16(data, offset);
                offset += 2;
                if (length == 0) continue;
                if (length < baseLength || length > data.Length - offset) {
                    throw new InvalidDataException("A Word 6/95 style record extends beyond the stylesheet.");
                }

                ReadOnlySpan<byte> record = data.Slice(offset, length);
                offset = (offset + length + 1) & ~1;
                ushort kindAndBase = LegacyWord69Reader.U16(record, 2);
                if ((kindAndBase & 15) != 1) continue;

                ushort nativeId = (ushort)(LegacyWord69Reader.U16(record, 0) & 0x0FFF);
                ushort baseIndex = (ushort)(kindAndBase >> 4);
                ushort nextAndCount = LegacyWord69Reader.U16(record, 4);
                ushort nextIndex = (ushort)(nextAndCount >> 4);
                int nameLength = baseLength < record.Length ? record[baseLength] : 0;
                if (nameLength == 0 || nameLength > record.Length - baseLength - 2
                    || record[baseLength + 1 + nameLength] != 0) {
                    properties.HasUnsupportedFormatting = true;
                    continue;
                }

                string name = LegacyWord69Properties.DecodeName(record.Slice(baseLength + 1, nameLength));
                int cursor = (baseLength + nameLength + 3) & ~1;
                var paragraph = LegacyDocParagraphFormat.Default;
                var character = LegacyDocCharacterFormat.Default;
                for (int propertyIndex = 0; propertyIndex < (nextAndCount & 15); propertyIndex++) {
                    if (cursor > record.Length - 2) throw new InvalidDataException("A Word 6/95 style property length is truncated.");
                    int propertyLength = LegacyWord69Reader.U16(record, cursor);
                    cursor += 2;
                    if (propertyLength > record.Length - cursor) throw new InvalidDataException("A Word 6/95 style property extends beyond its record.");
                    ReadOnlySpan<byte> property = record.Slice(cursor, propertyLength);
                    if (propertyIndex == 0 && propertyLength >= 2) {
                        paragraph = properties.ReadParagraph(property.Slice(2), LegacyWord69Reader.U16(property, 0));
                    } else if (propertyIndex == 1) {
                        character = properties.ReadCharacter(property, fonts, out int? fontIndex, out ushort? languageId);
                        if (fontIndex.HasValue) properties.StyleFonts[index] = fontIndex.Value;
                        if (languageId.HasValue) properties.StyleLanguages[index] = languageId.Value;
                    } else if (propertyLength != 0) {
                        properties.HasUnsupportedFormatting = true;
                    }

                    cursor = (cursor + propertyLength + 1) & ~1;
                }

                if (index == 0) character = character.WithDefaultFont(defaultFont);
                ushort? basedOn = baseIndex == 0x0FFF || baseIndex == index ? null : baseIndex;
                ushort? next = nextIndex == 0 ? null : nextIndex;
                styles[index] = TryMapBuiltInParagraphStyle(nativeId, name, out var builtIn)
                    ? LegacyDocParagraphStyle.ForBuiltIn(index, name, builtIn, basedOn, next, paragraph, character)
                    : LegacyDocParagraphStyle.ForCustom(index, name, CreateCustomStyleId(name, index, usedIds), basedOn, next, paragraph, character);
            }

            return styles.Count == 0 ? Empty : new LegacyDocStyleSheet(styles);
        }
    }
}

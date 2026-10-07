using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Reads MS-DOC PlfLst/LVL and PlfLfo/LFOData without trusting binary counts or offsets.</summary>
    internal static class LegacyDocNumberingReader {
        internal static LegacyDocNumbering Read(byte[] table, LegacyDocFib fib, IReadOnlyList<string> fonts, out string? warning) {
            warning = null;
            if (fib.LcbPlfLst == 0 && fib.LcbPlfLfo == 0) return LegacyDocNumbering.Empty;
            try {
                return ReadTables(table, fib, fonts);
            } catch (InvalidDataException exception) {
                warning = exception.Message;
                return LegacyDocNumbering.Empty;
            }
        }

        private static LegacyDocNumbering ReadTables(byte[] table, LegacyDocFib fib, IReadOnlyList<string> fonts) {
            RequireBounds(table, fib.FcPlfLst, fib.LcbPlfLst);
            RequireBounds(table, fib.FcPlfLfo, fib.LcbPlfLfo);
            if (fib.LcbPlfLst < 2 || fib.LcbPlfLfo < 4) throw Invalid("Missing native list definition or instance table.");
            int count = LegacyDocFib.ReadUInt16(table, fib.FcPlfLst);
            if (count > short.MaxValue || fib.LcbPlfLst != 2 + count * 28) throw Invalid("Invalid native list definition count.");
            int cursor = fib.FcPlfLst + fib.LcbPlfLst;
            int levelEnd = fib.FcPlfLfo >= cursor ? fib.FcPlfLfo : table.Length;
            var definitions = new List<LegacyDocListDefinition>(count);
            var ids = new HashSet<int>();
            var levelsById = new Dictionary<int, int>();
            for (int index = 0; index < count; index++) {
                int offset = fib.FcPlfLst + 2 + index * 28;
                int id = LegacyDocFib.ReadInt32(table, offset);
                byte flags = table[offset + 26];
                if (id == -1 || !ids.Add(id)) throw Invalid("Invalid or duplicate native list identifier.");
                if ((flags & 4) != 0) throw Invalid("Native automatic-number field list definitions are not projected.");
                int levelCount = (flags & 1) != 0 ? 1 : 9;
                levelsById.Add(id, levelCount);
                var levels = new List<LegacyDocListLevel>(levelCount);
                for (int level = 0; level < levelCount; level++) {
                    ushort style = LegacyDocFib.ReadUInt16(table, offset + 8 + level * 2);
                    levels.Add(ReadLevel(table, ref cursor, levelEnd, fonts, level, style == 0x0FFF ? null : (ushort?)style));
                }
                definitions.Add(new LegacyDocListDefinition(id, (flags & 16) != 0, levels));
            }
            int instanceCount = LegacyDocFib.ReadInt32(table, fib.FcPlfLfo);
            if (instanceCount < 0 || instanceCount > short.MaxValue || instanceCount > (fib.LcbPlfLfo - 4) / 20)
                throw Invalid("Invalid native list instance count.");
            int instanceEnd = fib.FcPlfLfo + fib.LcbPlfLfo;
            int data = fib.FcPlfLfo + 4 + instanceCount * 16;
            var instances = new List<LegacyDocListInstance>(instanceCount);
            for (int index = 0; index < instanceCount; index++) {
                int offset = fib.FcPlfLfo + 4 + index * 16;
                int id = LegacyDocFib.ReadInt32(table, offset);
                int overrideCount = table[offset + 12];
                if (!ids.Contains(id) || overrideCount > 9) throw Invalid("A native list instance has no valid definition.");
                if (data > instanceEnd - 4) throw Invalid("Truncated native list instance data.");
                data += 4; // LFOData.cp is advisory; paragraph iLfo identifies the instance.
                var overrides = new List<LegacyDocListOverride>(overrideCount);
                var overriddenLevels = new HashSet<int>();
                for (int item = 0; item < overrideCount; item++) {
                    if (data > instanceEnd - 8) throw Invalid("Truncated native list level override.");
                    int start = LegacyDocFib.ReadInt32(table, data);
                    byte flags = table[data + 4];
                    int level = flags & 15;
                    if (level >= levelsById[id] || !overriddenLevels.Add(level)) throw Invalid("Invalid or duplicate native list override level.");
                    bool hasStart = (flags & 16) != 0, hasFormatting = (flags & 32) != 0;
                    data += 8;
                    LegacyDocListLevel? formatting = hasFormatting ? ReadLevel(table, ref data, instanceEnd, fonts, level, null) : null;
                    int? effectiveStart = hasStart ? (hasFormatting ? formatting!.Start : start) : (int?)null;
                    if (effectiveStart < 0 || effectiveStart > 32767) throw Invalid("Native list start is outside its valid range.");
                    overrides.Add(new LegacyDocListOverride(level, effectiveStart, formatting));
                }
                instances.Add(new LegacyDocListInstance(id, overrides));
            }
            if (data != instanceEnd) throw Invalid("Native list instance data has an unexpected length.");
            return new LegacyDocNumbering(definitions, instances);
        }

        private static LegacyDocListLevel ReadLevel(byte[] table, ref int cursor, int end, IReadOnlyList<string> fonts, int level, ushort? style) {
            if (cursor < 0 || cursor > end - 28) throw Invalid("Truncated native list level header.");
            int start = LegacyDocFib.ReadInt32(table, cursor);
            byte nfc = table[cursor + 4], flags = table[cursor + 5], follow = table[cursor + 15], restart = table[cursor + 26];
            NumberFormatValues? format = nfc == 23 ? NumberFormatValues.Bullet : nfc == 255 ? NumberFormatValues.None : LegacyDocNumberFormatMapper.FromNfc(nfc);
            if (start < 0 || start > 32767 || format == null || nfc == 8 || nfc == 9 || nfc == 15 || nfc == 19)
                throw Invalid("Unsupported native list numbering format or start.");
            if ((flags & 3) > 2 || follow > 2 || ((flags & 8) != 0 && restart > level)) throw Invalid("Invalid native list level alignment, suffix or restart.");
            int chpxLength = table[cursor + 24], papxLength = table[cursor + 25];
            int papx = cursor + 28, chpx = papx + papxLength, textOffset = chpx + chpxLength;
            if (textOffset > end - 2) throw Invalid("Truncated native list level formatting.");
            int textLength = LegacyDocFib.ReadUInt16(table, textOffset);
            if (textLength > (end - textOffset - 2) / 2) throw Invalid("Truncated native list marker text.");
            string rawText = Encoding.Unicode.GetString(table, textOffset + 2, textLength * 2);
            var text = new StringBuilder(rawText.Length);
            var placeholderOffsets = new HashSet<int>();
            int previousPosition = -1;
            for (int index = 0; index < 9 && table[cursor + 6 + index] != 0; index++) {
                int position = table[cursor + 6 + index] - 1;
                if (position <= previousPosition || position >= rawText.Length || !placeholderOffsets.Add(position) || rawText[position] > level)
                    throw Invalid("Invalid native list level placeholder.");
                previousPosition = position;
            }
            if (nfc == 23 && (rawText.Length != 1 || placeholderOffsets.Count != 0)) throw Invalid("Invalid native list bullet text.");
            for (int index = 0; index < rawText.Length; index++) {
                char character = rawText[index];
                if (placeholderOffsets.Contains(index)) text.Append('%').Append((int)character + 1);
                else {
                    if (character < 32) throw Invalid("Unsupported control character in native list marker text.");
                    text.Append(character);
                }
            }
            var paragraph = LegacyDocParagraphFormattingReader.ReadGrpprl(table, papx, papxLength);
            var characterFormat = LegacyDocCharacterFormattingReader.ReadGrpprl(table, chpx, chpxLength, fonts);
            cursor = textOffset + 2 + textLength * 2;
            return new LegacyDocListLevel(start, format.Value, text.ToString(), flags, follow, restart, style, paragraph, characterFormat);
        }

        private static void RequireBounds(byte[] bytes, int offset, int length) {
            if (offset < 0 || length < 0 || offset > bytes.Length - length) throw Invalid("Native list table points outside the selected table stream.");
        }

        private static InvalidDataException Invalid(string message) => new InvalidDataException(message);
    }
}

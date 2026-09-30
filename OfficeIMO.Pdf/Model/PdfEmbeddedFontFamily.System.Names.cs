namespace OfficeIMO.Pdf;

public sealed partial class PdfEmbeddedFontFamily {
    internal const int MaxSystemFontNameAliases = 128;
    internal const int MaxSystemFontNameAliasCharacters = 65536;

    private static bool TryReadTrueTypeNameMetadata(byte[] data, out TrueTypeNameMetadata? metadata) {
        return TryReadTrueTypeNameMetadata(data, out metadata, out _);
    }

    private static bool TryReadTrueTypeNameMetadata(byte[] data, out TrueTypeNameMetadata? metadata, out bool nameTableAbsent) {
        metadata = null;
        nameTableAbsent = false;
        try {
            if (data.Length < 12) {
                return false;
            }

            var tables = ReadFontTableDirectory(data);
            if (!tables.TryGetValue("name", out FontTableRecord nameTable)) {
                nameTableAbsent = true;
                return false;
            }

            return TryReadTrueTypeNameTable(data, nameTable.Offset,
                nameTable.Length, out metadata);
        } catch (System.Exception exception) when (exception is System.NotSupportedException) {
            return false;
        }
    }

    private static bool TryReadTrueTypeNameTable(byte[] data, int offset,
        int tableLength, out TrueTypeNameMetadata? metadata) {
        metadata = null;
        try {
            EnsureRange(data, offset, tableLength);
            var names = new System.Collections.Generic.Dictionary<int, TrueTypeNameValue>();
            var aliases = new System.Collections.Generic.HashSet<string>(System.StringComparer.Ordinal);
            int aliasCharacters = 0;
            var familyAliases = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            var typographicFamilyAliases = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            int count = ReadUInt16(data, offset + 2);
            if (count > MaxSystemFontNameRecords || 6L + count * 12L > tableLength) return false;
            int decodedNameBytes = 0;
            int stringOffset = offset + ReadUInt16(data, offset + 4);
            for (int i = 0; i < count; i++) {
                int record = offset + 6 + i * 12;
                EnsureRange(data, record, 12);
                int platformId = ReadUInt16(data, record);
                int encodingId = ReadUInt16(data, record + 2);
                int languageId = ReadUInt16(data, record + 4);
                int nameId = ReadUInt16(data, record + 6);
                if (nameId != 1 && nameId != 2 && nameId != 4 && nameId != 6 && nameId != 16 && nameId != 17) {
                    continue;
                }

                int length = ReadUInt16(data, record + 8);
                int valueOffset = stringOffset + ReadUInt16(data, record + 10);
                EnsureRange(data, valueOffset, length);
                if (length > MaxSystemFontDecodedNameBytes - decodedNameBytes) return false;
                decodedNameBytes += length;
                string? value = DecodeNameValue(data, valueOffset, length, platformId, encodingId);
                if (string.IsNullOrWhiteSpace(value)) {
                    continue;
                }

                string trimmed = value!.Trim();
                if ((nameId == 1 || nameId == 4 || nameId == 6 || nameId == 16) && aliases.Add(trimmed)) {
                    aliasCharacters = checked(aliasCharacters + trimmed.Length);
                    if (aliases.Count > MaxSystemFontNameAliases || aliasCharacters > MaxSystemFontNameAliasCharacters) {
                        throw new System.NotSupportedException("TrueType font name aliases exceed the retained metadata budget.");
                    }
                }

                if (nameId == 1) familyAliases.Add(trimmed);
                if (nameId == 16) typographicFamilyAliases.Add(trimmed);
                int score = GetNameValueScore(platformId, languageId);
                if (!names.TryGetValue(nameId, out TrueTypeNameValue? existing) || score > existing.Score) {
                    names[nameId] = new TrueTypeNameValue(trimmed, score);
                }
            }

            metadata = new TrueTypeNameMetadata(
                GetName(names, 1),
                GetName(names, 2),
                GetName(names, 4),
                GetName(names, 6),
                GetName(names, 16),
                GetName(names, 17), aliases, familyAliases, typographicFamilyAliases);
            return true;
        } catch (System.Exception exception) when (exception is System.NotSupportedException) {
            return false;
        }
    }

    private static string? DecodeNameValue(byte[] data, int offset, int length, int platformId, int encodingId) {
        if (platformId == 3 || platformId == 0) {
            return length % 2 == 0
                ? System.Text.Encoding.BigEndianUnicode.GetString(data, offset, length).TrimEnd('\0')
                : null;
        }

        if (platformId == 1 && encodingId == 0) {
            return System.Text.Encoding.ASCII.GetString(data, offset, length).TrimEnd('\0');
        }

        return null;
    }

    private static int GetNameValueScore(int platformId, int languageId) {
        if (platformId == 3) {
            return languageId == 0x0409 ? 50 : (languageId & 0x03ff) == 0x0009 ? 40 : 20;
        }

        if (platformId == 0) {
            return 30;
        }

        return 10;
    }

    private static string? GetName(System.Collections.Generic.Dictionary<int, TrueTypeNameValue> names, int nameId) =>
        names.TryGetValue(nameId, out TrueTypeNameValue? value) ? value.Value : null;

    private sealed class TrueTypeNameMetadata {
        public TrueTypeNameMetadata(
            string? familyName,
            string? subfamilyName,
            string? fullName,
            string? postScriptName,
            string? typographicFamilyName,
            string? typographicSubfamilyName,
            System.Collections.Generic.IReadOnlyCollection<string> aliases,
            System.Collections.Generic.IReadOnlyCollection<string> familyAliases,
            System.Collections.Generic.IReadOnlyCollection<string> typographicFamilyAliases) {
            FamilyName = familyName;
            SubfamilyName = subfamilyName;
            FullName = fullName;
            PostScriptName = postScriptName;
            TypographicFamilyName = typographicFamilyName;
            TypographicSubfamilyName = typographicSubfamilyName;
            Aliases = aliases;
            FamilyAliases = familyAliases;
            TypographicFamilyAliases = typographicFamilyAliases;
        }

        public System.Collections.Generic.IReadOnlyCollection<string> Aliases { get; }
        public System.Collections.Generic.IReadOnlyCollection<string> FamilyAliases { get; }
        public System.Collections.Generic.IReadOnlyCollection<string> TypographicFamilyAliases { get; }

        public string? FamilyName { get; }

        public string? SubfamilyName { get; }

        public string? FullName { get; }

        public string? PostScriptName { get; }

        public string? TypographicFamilyName { get; }

        public string? TypographicSubfamilyName { get; }

        public System.Collections.Generic.IEnumerable<string?> GetFamilyNames() {
            yield return TypographicFamilyName;
            yield return FamilyName;
            // Keep alternative language/platform names discoverable without changing face classification.
            foreach (string alias in Aliases) yield return alias;
        }

        public System.Collections.Generic.IEnumerable<string?> GetFaceNames() {
            yield return FullName;
            yield return PostScriptName;
            yield return CombineFamilyAndSubfamily(TypographicFamilyName, TypographicSubfamilyName);
            yield return CombineFamilyAndSubfamily(FamilyName, SubfamilyName);
        }

        private static string? CombineFamilyAndSubfamily(string? familyName, string? subfamilyName) {
            if (string.IsNullOrWhiteSpace(familyName)) {
                return null;
            }

            if (string.IsNullOrWhiteSpace(subfamilyName) ||
                string.Equals(subfamilyName, "Regular", System.StringComparison.OrdinalIgnoreCase)) {
                return familyName;
            }

            return familyName + " " + subfamilyName;
        }
    }

    private sealed class TrueTypeNameValue {
        public TrueTypeNameValue(string value, int score) {
            Value = value;
            Score = score;
        }

        public string Value { get; }

        public int Score { get; }
    }

}

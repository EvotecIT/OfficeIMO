using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Word;

/// <summary>
/// Isolates RTF list identifiers before Word merges an AltChunk into the document.
/// </summary>
/// <remarks>
/// Independently produced RTF documents commonly reuse the same list, list-override,
/// and template identifiers. Word can then bind a later fragment to an earlier list
/// definition, for example rendering bullets as the next values of a numbered list.
/// This normalizer preserves identifiers when they are unused and remaps only values
/// that collide with an RTF fragment already embedded in the same document.
/// </remarks>
internal static class RtfListSemanticsNormalizer {
    private const string RtfContentType = "application/rtf";
    private const int MaximumExistingRtfPartCount = 1_024;
    private const long MaximumExistingRtfPartBytes = 16L * 1024L * 1024L;
    private const long MaximumExistingRtfTotalBytes = 64L * 1024L * 1024L;
    private static readonly ConditionalWeakTable<MainDocumentPart, OccupiedIdentifierState> IdentifierStates =
        new ConditionalWeakTable<MainDocumentPart, OccupiedIdentifierState>();

    internal static string Normalize(string rtf, MainDocumentPart mainDocumentPart) {
        if (string.IsNullOrEmpty(rtf)
            || (rtf.IndexOf("\\list", StringComparison.OrdinalIgnoreCase) < 0
                && rtf.IndexOf("\\ls", StringComparison.OrdinalIgnoreCase) < 0)) {
            return rtf;
        }

        var current = new IdentifierSets();
        CollectIdentifiers(rtf, current);
        OccupiedIdentifierState state = IdentifierStates.GetValue(
            mainDocumentPart,
            _ => new OccupiedIdentifierState());
        lock (state.SyncRoot) {
            if (!state.Initialized) {
                CollectExistingIdentifiers(mainDocumentPart, state.Identifiers);
                state.Initialized = true;
            }

            string normalized = rtf;
            if (current.Overlaps(state.Identifiers)) {
                var mappings = IdentifierMappings.Create(current, state.Identifiers);
                normalized = RewriteIdentifiers(rtf, mappings);
                current = new IdentifierSets();
                CollectIdentifiers(normalized, current);
            }

            state.Identifiers.UnionWith(current);
            return normalized;
        }
    }

    private static void CollectExistingIdentifiers(
        MainDocumentPart mainDocumentPart,
        IdentifierSets identifiers) {
        int partCount = 0;
        long totalBytes = 0;
        foreach (AlternativeFormatImportPart part in mainDocumentPart.AlternativeFormatImportParts) {
            if (!string.Equals(part.ContentType, RtfContentType, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            partCount++;
            if (partCount > MaximumExistingRtfPartCount) {
                throw new InvalidDataException($"RTF altChunk count exceeds the configured limit of {MaximumExistingRtfPartCount}.");
            }

            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            long remaining = MaximumExistingRtfTotalBytes - totalBytes;
            if (remaining <= 0) {
                throw new InvalidDataException($"RTF altChunk bytes exceed the configured aggregate limit of {MaximumExistingRtfTotalBytes}.");
            }
            byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, Math.Min(MaximumExistingRtfPartBytes, remaining));
            totalBytes = checked(totalBytes + bytes.LongLength);
            using var reader = new StreamReader(
                new MemoryStream(bytes, writable: false),
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true);
            CollectIdentifiers(reader.ReadToEnd(), identifiers);
        }
    }

    private static void CollectIdentifiers(string rtf, IdentifierSets identifiers) {
        ScanControlWords(rtf, (kind, _, _, value) => identifiers.Get(kind).Add(value));
    }

    private static string RewriteIdentifiers(string rtf, IdentifierMappings mappings) {
        var replacements = new List<Replacement>();
        ScanControlWords(rtf, (kind, start, length, value) => {
            int replacement = mappings.Get(kind)[value];
            if (replacement != value) {
                replacements.Add(new Replacement(start, length, replacement));
            }
        });

        if (replacements.Count == 0) {
            return rtf;
        }

        var normalized = new StringBuilder(rtf);
        for (int index = replacements.Count - 1; index >= 0; index--) {
            Replacement replacement = replacements[index];
            normalized.Remove(replacement.Start, replacement.Length);
            normalized.Insert(replacement.Start, replacement.Value.ToString(CultureInfo.InvariantCulture));
        }
        return normalized.ToString();
    }

    private static void ScanControlWords(string rtf, Action<IdentifierKind, int, int, int> visitor) {
        int position = 0;
        while (position < rtf.Length) {
            if (rtf[position] != '\\' || position + 1 >= rtf.Length) {
                position++;
                continue;
            }

            int nameStart = position + 1;
            if (!char.IsLetter(rtf[nameStart])) {
                position += 2;
                continue;
            }

            int nameEnd = nameStart + 1;
            while (nameEnd < rtf.Length && char.IsLetter(rtf[nameEnd])) {
                nameEnd++;
            }

            int numberStart = nameEnd;
            int numberEnd = numberStart;
            if (numberEnd < rtf.Length && rtf[numberEnd] == '-') {
                numberEnd++;
            }
            int digitStart = numberEnd;
            while (numberEnd < rtf.Length && char.IsDigit(rtf[numberEnd])) {
                numberEnd++;
            }

            int value = 0;
            bool hasParameter = numberEnd > digitStart
                && int.TryParse(
                    rtf.Substring(numberStart, numberEnd - numberStart),
                    NumberStyles.AllowLeadingSign,
                    CultureInfo.InvariantCulture,
                    out value);

            if (ControlWordEquals(rtf, nameStart, nameEnd - nameStart, "bin")) {
                int binaryStart = numberEnd;
                if (binaryStart < rtf.Length && rtf[binaryStart] == ' ') {
                    binaryStart++;
                }
                if (hasParameter && value >= 0) {
                    position = (int)Math.Min(rtf.Length, (long)binaryStart + value);
                } else {
                    position = binaryStart;
                }
                continue;
            }

            if (hasParameter
                && TryGetIdentifierKind(rtf, nameStart, nameEnd - nameStart, out IdentifierKind kind)) {
                visitor(kind, numberStart, numberEnd - numberStart, value);
            }
            position = numberEnd;
        }
    }

    private static bool ControlWordEquals(
        string rtf,
        int start,
        int length,
        string expected) {
        return length == expected.Length
            && string.Compare(rtf, start, expected, 0, length, StringComparison.OrdinalIgnoreCase) == 0;
    }

    private static bool TryGetIdentifierKind(
        string rtf,
        int start,
        int length,
        out IdentifierKind kind) {
        if (ControlWordEquals(rtf, start, length, "ls")) {
            kind = IdentifierKind.ListOverride;
            return true;
        }
        if (ControlWordEquals(rtf, start, length, "listid")) {
            kind = IdentifierKind.List;
            return true;
        }
        if (ControlWordEquals(rtf, start, length, "listtemplateid")) {
            kind = IdentifierKind.Template;
            return true;
        }

        kind = default;
        return false;
    }

    private enum IdentifierKind {
        List,
        ListOverride,
        Template
    }

    private readonly struct Replacement {
        internal Replacement(int start, int length, int value) {
            Start = start;
            Length = length;
            Value = value;
        }

        internal int Start { get; }
        internal int Length { get; }
        internal int Value { get; }
    }

    private sealed class IdentifierSets {
        internal HashSet<int> Lists { get; } = new HashSet<int>();
        internal HashSet<int> ListOverrides { get; } = new HashSet<int>();
        internal HashSet<int> Templates { get; } = new HashSet<int>();

        internal HashSet<int> Get(IdentifierKind kind) {
            return kind switch {
                IdentifierKind.List => Lists,
                IdentifierKind.ListOverride => ListOverrides,
                IdentifierKind.Template => Templates,
                _ => throw new ArgumentOutOfRangeException(nameof(kind))
            };
        }

        internal bool Overlaps(IdentifierSets other) {
            return Lists.Overlaps(other.Lists)
                || ListOverrides.Overlaps(other.ListOverrides)
                || Templates.Overlaps(other.Templates);
        }

        internal void UnionWith(IdentifierSets other) {
            Lists.UnionWith(other.Lists);
            ListOverrides.UnionWith(other.ListOverrides);
            Templates.UnionWith(other.Templates);
        }
    }

    private sealed class OccupiedIdentifierState {
        internal object SyncRoot { get; } = new object();
        internal IdentifierSets Identifiers { get; } = new IdentifierSets();
        internal bool Initialized { get; set; }
    }

    private sealed class IdentifierMappings {
        private IdentifierMappings() { }

        internal Dictionary<int, int> Lists { get; } = new Dictionary<int, int>();
        internal Dictionary<int, int> ListOverrides { get; } = new Dictionary<int, int>();
        internal Dictionary<int, int> Templates { get; } = new Dictionary<int, int>();

        internal Dictionary<int, int> Get(IdentifierKind kind) {
            return kind switch {
                IdentifierKind.List => Lists,
                IdentifierKind.ListOverride => ListOverrides,
                IdentifierKind.Template => Templates,
                _ => throw new ArgumentOutOfRangeException(nameof(kind))
            };
        }

        internal static IdentifierMappings Create(IdentifierSets current, IdentifierSets occupied) {
            var mappings = new IdentifierMappings();
            CreateMap(current.Lists, occupied.Lists, mappings.Lists);
            CreateMap(current.ListOverrides, occupied.ListOverrides, mappings.ListOverrides);
            CreateMap(current.Templates, occupied.Templates, mappings.Templates);
            return mappings;
        }

        private static void CreateMap(
            HashSet<int> current,
            HashSet<int> occupied,
            Dictionary<int, int> mappings) {
            var reserved = new HashSet<int>(occupied);
            foreach (int value in current.OrderBy(value => value)) {
                if (!occupied.Contains(value)) {
                    mappings.Add(value, value);
                    reserved.Add(value);
                }
            }

            foreach (int value in current.OrderBy(value => value)) {
                if (!occupied.Contains(value)) {
                    continue;
                }

                int replacement = FindAvailableIdentifier(reserved);
                mappings.Add(value, replacement);
                reserved.Add(replacement);
            }
        }

        private static int FindAvailableIdentifier(HashSet<int> reserved) {
            for (int candidate = 1; candidate < int.MaxValue; candidate++) {
                if (!reserved.Contains(candidate)) {
                    return candidate;
                }
            }
            throw new InvalidOperationException("No RTF list identifier is available.");
        }
    }
}

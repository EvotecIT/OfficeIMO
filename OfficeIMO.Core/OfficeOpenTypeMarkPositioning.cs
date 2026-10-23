using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Describes an attachment without assuming a logical run's eventual visual pen direction.</summary>
internal readonly struct OfficeOpenTypeMarkAttachment {
    internal OfficeOpenTypeMarkAttachment(int parentGlyphIndex, int deltaX, int deltaY, bool isMark, bool isFontDirectedNoOp = false) {
        ParentGlyphIndex = parentGlyphIndex; DeltaX = deltaX; DeltaY = deltaY; IsMark = isMark;
        IsFontDirectedNoOp = isFontDirectedNoOp;
    }
    internal int ParentGlyphIndex { get; }
    internal int DeltaX { get; }
    internal int DeltaY { get; }
    internal bool IsMark { get; }
    /// <summary>The font selected a supported mark lookup, but supplied no applicable parent anchor.</summary>
    internal bool IsFontDirectedNoOp { get; }
}

/// <summary>Bounded static GPOS mark-to-base, mark-to-ligature and mark-to-mark attachment reader.</summary>
/// <remarks>Results retain logical indexes. The shaping owner resolves parent positions after bidi reordering.</remarks>
internal sealed partial class OfficeOpenTypeMarkPositioning {
    private readonly byte[] _data;
    private readonly int _start, _end, _scriptList, _featureList, _lookupList;
    private readonly OfficeOpenTypeGlyphFilter _filter;
    private readonly CancellationToken _cancellationToken;
    private int _operations;
    private const int MaximumOperations = 1_000_000;
    private const int MaximumRecords = 4096;

    private OfficeOpenTypeMarkPositioning(byte[] data, OfficeOpenTypeReader reader,
        OfficeOpenTypeGlyphFilter filter, CancellationToken cancellationToken) {
        _data = data; _filter = filter; _cancellationToken = cancellationToken;
        if (!reader.TryGetTable("GPOS", out _start, out int length)) { _start = _end = -1; return; }
        _end = checked(_start + length);
        Ensure(_start, 10);
        if (U16(_start) != 1 || U16(_start + 2) > 1) throw new InvalidDataException("Unsupported GPOS version.");
        if (U16(_start + 2) == 1) {
            Ensure(_start, 14);
            if (U32(_start + 10) != 0) throw new InvalidDataException("GPOS feature variations require a variable-layout provider.");
        }
        _scriptList = Relative(_start, U16(_start + 4), 2);
        _featureList = Relative(_start, U16(_start + 6), 2);
        _lookupList = Relative(_start, U16(_start + 8), 2);
    }

    /// <summary>Accepts interpreted mark positioning, including a supported font lookup's deliberate no-op.</summary>
    /// <remarks>Uncovered marks and unsupported layout remain declined; missing font anchors are never invented.</remarks>
    internal static bool TryGetAttachments(byte[] fontData, IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> tokens,
        string? language, OfficeTextFeatureSettings settings, CancellationToken cancellationToken,
        out OfficeOpenTypeMarkAttachment[] attachments) {
        if (fontData == null) throw new ArgumentNullException(nameof(fontData));
        if (tokens == null) throw new ArgumentNullException(nameof(tokens));
        if (settings == null) throw new ArgumentNullException(nameof(settings));
        cancellationToken.ThrowIfCancellationRequested();
        attachments = new OfficeOpenTypeMarkAttachment[tokens.Count];
        try {
            OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(fontData);
            OfficeOpenTypeGlyphFilter? filter = OfficeOpenTypeGlyphFilter.TryCreate(fontData);
            if (reader == null || filter == null) return false;
            var owner = new OfficeOpenTypeMarkPositioning(fontData, reader, filter, cancellationToken);
            return owner.Position(tokens, language, settings, attachments);
        } catch (Exception exception) when (exception is InvalidDataException or OverflowException
            or ArgumentOutOfRangeException or IndexOutOfRangeException) {
            attachments = Array.Empty<OfficeOpenTypeMarkAttachment>();
            return false;
        }
    }

    private bool Position(IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> tokens, string? language,
        OfficeTextFeatureSettings settings, OfficeOpenTypeMarkAttachment[] result) {
        bool hasMarks = false;
        var scripts = new string[tokens.Count];
        string inherited = "DFLT";
        for (int index = 0; index < tokens.Count; index++) {
            Charge();
            OfficeOpenTypeSubstitution.GlyphToken token = tokens[index];
            // CGJ is a Unicode nonspacing mark, but its retained, nonpainting
            // source token has no anchor. Normalization owns its ordering barrier.
            bool mark = !OfficeManagedTextShaper.IsNonPaintingControl(token.Scalar)
                && (OfficeOpenTypeGlyphFilter.IsUnicodeMark(token.Scalar)
                    || _filter.GetGlyphClass(token.GlyphId, token.Scalar, token.LigatureComponentTextIndexes?.Count > 1, Charge) == 3);
            hasMarks |= mark;
            result[index] = new OfficeOpenTypeMarkAttachment(-1, 0, 0, mark);
            string script = Script(token.Scalar);
            if (script != "DFLT" && !mark) inherited = script;
            scripts[index] = mark ? inherited : script;
        }
        // Leading inherited marks belong to the next supported script, if any.
        string following = "DFLT";
        for (int index = tokens.Count - 1; index >= 0; index--) {
            Charge();
            if (scripts[index] != "DFLT") following = scripts[index];
            else if (result[index].IsMark) scripts[index] = following;
        }
        bool marksEnabled = Enabled(settings, "mark");
        if (_start < 0) return !hasMarks || !marksEnabled;
        var interpretedMarks = new bool[tokens.Count];
        var scriptSet = new HashSet<string>(scripts, StringComparer.Ordinal);
        // Cursive attachment changes the origins of base letters and must not be silently omitted.
        foreach (string script in scriptSet) {
            List<int> cursive = Lookups(script, language, "curs", settings);
            if (Enabled(settings, "curs") && cursive.Count != 0) return false;
        }
        var active = new SortedDictionary<int, HashSet<string>>();
        foreach (string feature in new[] { "mark", "mkmk" }) {
            if (!Enabled(settings, feature)) continue;
            foreach (string script in scriptSet) {
                foreach (int lookupIndex in Lookups(script, language, feature, settings)) {
                    if (!active.TryGetValue(lookupIndex, out HashSet<string>? selected)) {
                        selected = new HashSet<string>(StringComparer.Ordinal); active.Add(lookupIndex, selected);
                    }
                    selected.Add(script);
                }
            }
        }
        // Both features participate in one GPOS stage. Their shared LookupList
        // indexes determine execution order, including repeated references.
        foreach (KeyValuePair<int, HashSet<string>> lookup in active)
            ApplyLookup(tokens, scripts, result, interpretedMarks, lookup.Key, lookup.Value);
        if (marksEnabled) {
            for (int index = 0; index < result.Length; index++) {
                Charge();
                if (!result[index].IsMark || result[index].ParentGlyphIndex >= 0) continue;
                if (!interpretedMarks[index]) return false;
                // Applying an understood font program can legitimately do nothing:
                // Source Sans Pro covers acutecomb but not precomposed a-dotbelow
                // as its parent. Native OpenType shaping retains the zero-advance
                // mark's nominal position; this is not an unavailable lookup.
                result[index] = new OfficeOpenTypeMarkAttachment(-1, 0, 0, isMark: true, isFontDirectedNoOp: true);
            }
        }
        return true;
    }

    private List<int> Lookups(string scriptTag, string? language, string featureTag,
        OfficeTextFeatureSettings settings) {
        int scriptCount = Count(_scriptList, 6), script = 0, fallback = 0;
        for (int index = 0; index < scriptCount; index++) {
            Charge(); int record = _scriptList + 2 + index * 6;
            if (Tag(record, scriptTag)) script = Relative(_scriptList, U16(record + 4), 4);
            if (Tag(record, "DFLT")) fallback = Relative(_scriptList, U16(record + 4), 4);
        }
        if (script == 0) script = fallback;
        var result = new List<int>();
        if (script == 0) return result;
        int languageSystem = Optional(script, U16(script), 6);
        string? languageTag = LanguageTag(language);
        int languages = U16(script + 2);
        if (languages > MaximumRecords) throw new InvalidDataException("GPOS language count exceeds the managed limit.");
        Ensure(script + 4, checked(languages * 6));
        for (int index = 0; languageTag != null && index < languages; index++) {
            Charge(); int record = script + 4 + index * 6;
            if (Tag(record, languageTag)) { languageSystem = Relative(script, U16(record + 4), 6); break; }
        }
        if (languageSystem == 0) return result;
        int featureCount = Count(_featureList, 6), selectedCount = U16(languageSystem + 4);
        if (selectedCount > MaximumRecords) throw new InvalidDataException("GPOS language features exceed the managed limit.");
        Ensure(languageSystem + 6, checked(selectedCount * 2));
        var selected = new List<int>();
        int required = U16(languageSystem + 2);
        if (required != ushort.MaxValue) selected.Add(required);
        for (int index = 0; index < selectedCount; index++) selected.Add(U16(languageSystem + 6 + index * 2));
        foreach (int index in selected) {
            Charge();
            if (index >= featureCount) throw new InvalidDataException("GPOS feature index is invalid.");
            int record = _featureList + 2 + index * 6;
            if (index == required) {
                string requiredTag = new(new[] { (char)_data[record], (char)_data[record + 1],
                    (char)_data[record + 2], (char)_data[record + 3] });
                if (settings.TryGetValue(requiredTag, out int requiredSetting) && requiredSetting == 0)
                    throw new InvalidDataException("Disabling a required GPOS feature cannot preserve its layout contract.");
            }
            bool supported = Tag(record, "kern") || Tag(record, "mark") || Tag(record, "mkmk");
            if (!supported) {
                string tag = new(new[] { (char)_data[record], (char)_data[record + 1],
                    (char)_data[record + 2], (char)_data[record + 3] });
                if (index == required || settings.TryGetValue(tag, out int authored) && authored > 0)
                    throw new InvalidDataException("The selected GPOS feature requires an unsupported positioning composition.");
            }
            if (!Tag(record, featureTag)) continue;
            int feature = Relative(_featureList, U16(record + 4), 4), count = U16(feature + 2);
            if (count > MaximumRecords) throw new InvalidDataException("GPOS feature lookups exceed the managed limit.");
            Ensure(feature + 4, checked(count * 2));
            for (int lookup = 0; lookup < count; lookup++) { Charge(); result.Add(U16(feature + 4 + lookup * 2)); }
        }
        return result;
    }

    private static bool Enabled(OfficeTextFeatureSettings settings, string tag) => !settings.TryGetValue(tag, out int value) || value != 0;
    private static string Script(int scalar) => scalar is >= 0x0590 and <= 0x05FF ? "hebr"
        : scalar is >= 0x0600 and <= 0x06FF or >= 0x0750 and <= 0x077F or >= 0x08A0 and <= 0x08FF ? "arab"
        : scalar is >= 0x0041 and <= 0x024F or >= 0x1E00 and <= 0x1EFF ? "latn" : "DFLT";
    private static string? LanguageTag(string? language) {
        string? primary = language?.Split('-')[0].ToLowerInvariant();
        return primary switch { "ar" => "ARA ", "fa" => "FAR ", "ur" => "URD ", "he" or "iw" => "IWR ",
            "en" => "ENG ", "de" => "DEU ", "fr" => "FRA ", "tr" => "TRK ", _ => null };
    }

    private void Charge() {
        if (++_operations > MaximumOperations) throw new InvalidDataException("OpenType mark positioning exceeded the managed work budget.");
        if ((_operations & 255) == 0) _cancellationToken.ThrowIfCancellationRequested();
    }
    private int Count(int table, int stride) {
        int count = U16(table);
        if (count > MaximumRecords) throw new InvalidDataException("GPOS record count exceeds the managed limit.");
        Ensure(table + 2, checked(count * stride)); return count;
    }
    private int Optional(int table, int relative, int size) => relative == 0 ? 0 : Relative(table, (uint)relative, size);
    private int Relative(int table, uint relative, int size) {
        if (relative == 0 || relative > int.MaxValue) throw new InvalidDataException("Invalid GPOS relative offset.");
        int result = checked(table + (int)relative); Ensure(result, size); return result;
    }
    private void Ensure(int offset, int size) {
        if (size < 0 || offset < _start || offset > _end - size) throw new InvalidDataException("OpenType GPOS data is truncated.");
    }
    private bool Tag(int offset, string tag) { Ensure(offset, 4); return _data[offset] == tag[0] && _data[offset + 1] == tag[1] && _data[offset + 2] == tag[2] && _data[offset + 3] == tag[3]; }
    private ushort U16(int offset) { Ensure(offset, 2); return (ushort)((_data[offset] << 8) | _data[offset + 1]); }
    private short S16(int offset) => unchecked((short)U16(offset));
    private uint U32(int offset) { Ensure(offset, 4); return ((uint)_data[offset] << 24) | ((uint)_data[offset + 1] << 16) | ((uint)_data[offset + 2] << 8) | _data[offset + 3]; }
}

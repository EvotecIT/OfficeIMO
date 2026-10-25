using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeKerning {
    /// <summary>
    /// Applies default kern lookups before visual reordering. Context and ignored glyphs are
    /// interpreted in logical order; the caller carries each result with its source token.
    /// </summary>
    /// <remarks>Unsupported active positioning declines the run instead of dropping its adjustments.</remarks>
    internal bool TryPositionRun(IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> logicalTokens,
        string? language, CancellationToken cancellationToken, out OfficeOpenTypeGlyphPositioning[] positions,
        bool rejectMarkAdjustments = false) {
        if (logicalTokens == null) throw new ArgumentNullException(nameof(logicalTokens));
        cancellationToken.ThrowIfCancellationRequested();
        positions = new OfficeOpenTypeGlyphPositioning[logicalTokens.Count];
        try {
            OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(_data);
            OfficeOpenTypeGlyphFilter? filter = OfficeOpenTypeGlyphFilter.TryCreate(_data);
            if (reader == null || filter == null) return false;
            var run = new LogicalPositioningRun(this, reader, filter, logicalTokens, language,
                cancellationToken, positions, rejectMarkAdjustments);
            run.Apply();
            return true;
        } catch (Exception exception) when (exception is InvalidDataException or OverflowException
            or ArgumentOutOfRangeException or IndexOutOfRangeException) {
            positions = Array.Empty<OfficeOpenTypeGlyphPositioning>();
            return false;
        }
    }

    private int FindPositioningLanguageSystem(int script, string? language, PositioningWorkBudget budget) {
        if (!InPositioningBounds(script, 4, budget)) return -1;
        int relative = ReadUInt16(script);
        int selected = relative == 0 ? -1 : checked(script + relative);
        string? primary = language?.Split('-')[0].ToLowerInvariant();
        string? tag = primary switch {
            "ar" => "ARA ", "fa" => "FAR ", "ur" => "URD ", "ps" => "PAS ",
            "sd" => "SND ", "ug" => "UYG ", "ku" => "KUR ", "he" or "iw" => "IWR ",
            "en" => "ENG ", "fr" => "FRA ", "de" => "DEU ", "tr" => "TRK ", _ => null
        };
        // Preserve the original default-language route for callers without a language hint.
        if (tag != null) {
            int count = ReadUInt16(script + 2);
            if (!InPositioningBounds(script + 4, checked(count * 6), budget)) return -1;
            for (int index = 0; index < count; index++) {
                budget.Charge();
                int record = script + 4 + index * 6;
                if (TagEquals(record, tag)) { selected = checked(script + ReadUInt16(record + 4)); break; }
            }
        }
        return selected >= 0 && InPositioningBounds(selected, 6, budget) ? selected : -1;
    }

    /// <summary>One bounded logical GPOS pass, sharing pair decoding with ordinary font consumers.</summary>
    private sealed partial class LogicalPositioningRun {
        private const int MaximumContextGlyphs = 64;
        private const int MaximumRecords = 4096;
        private const int MaximumNesting = 8;
        private readonly OfficeOpenTypeKerning _owner;
        private readonly OfficeOpenTypeGlyphFilter _filter;
        private readonly IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> _tokens;
        private readonly OfficeOpenTypeGlyphPositioning[] _positions;
        private readonly string[] _scripts;
        private readonly string? _language;
        private readonly int _start, _end, _scriptList, _featureList, _lookupList;
        private readonly PositioningWorkBudget _budget;
        private readonly bool _rejectMarkAdjustments;

        internal LogicalPositioningRun(OfficeOpenTypeKerning owner, OfficeOpenTypeReader reader,
            OfficeOpenTypeGlyphFilter filter, IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> tokens,
            string? language, CancellationToken cancellationToken, OfficeOpenTypeGlyphPositioning[] positions,
            bool rejectMarkAdjustments) {
            _owner = owner; _filter = filter; _tokens = tokens; _language = language; _positions = positions;
            _rejectMarkAdjustments = rejectMarkAdjustments;
            _scripts = ResolveScripts(tokens);
            if (!reader.TryGetTable("GPOS", out _start, out int length)) {
                _start = _end = -1;
                _budget = new PositioningWorkBudget(cancellationToken);
                return;
            }
            _end = checked(_start + length);
            _budget = new PositioningWorkBudget(cancellationToken, _start, _end);
            Ensure(_start, 10);
            if (_start != owner._gpos || U16(_start) != 1 || U16(_start + 2) > 1)
                throw new InvalidDataException("Unsupported GPOS header.");
            if (U16(_start + 2) == 1) {
                Ensure(_start, 14);
                if (U32(_start + 10) != 0)
                    throw new InvalidDataException("GPOS feature variations require a variable-layout provider.");
            }
            _scriptList = Relative(_start, U16(_start + 4), 2);
            _featureList = Relative(_start, U16(_start + 6), 2);
            _lookupList = Relative(_start, U16(_start + 8), 2);
            ValidateSelectionTables();
        }

        internal void Apply() {
            var byScript = new Dictionary<string, HashSet<ushort>>(StringComparer.Ordinal);
            var active = new SortedDictionary<ushort, HashSet<string>>();
            foreach (string script in new HashSet<string>(_scripts, StringComparer.Ordinal)) {
                var selected = new HashSet<ushort>();
                if (_start >= 0) {
                    foreach (ushort lookup in _owner.GposFeatureLookupIndexes(_scriptList, _featureList,
                        "kern", script, _budget, _language)) {
                        if (lookup >= U16(_lookupList)) throw new InvalidDataException("Invalid GPOS lookup index.");
                        selected.Add(lookup);
                        if (!active.TryGetValue(lookup, out HashSet<string>? scripts)) {
                            scripts = new HashSet<string>(StringComparer.Ordinal); active.Add(lookup, scripts);
                        }
                        scripts.Add(script);
                    }
                }
                byScript.Add(script, selected);
            }
            foreach (KeyValuePair<ushort, HashSet<string>> lookup in active) {
                for (int index = 0; index < _tokens.Count;) {
                    _budget.Charge();
                    if (!lookup.Value.Contains(_scripts[index])) { index++; continue; }
                    bool matched = ApplyLookup(lookup.Key, index, _tokens.Count - 1, 0, out int next);
                    index = matched ? Math.Max(index + 1, next) : index + 1;
                }
            }
            for (int index = 0; index + 1 < _tokens.Count; index++) {
                _budget.Charge();
                // A selected GPOS kern feature supersedes the legacy table even when a
                // particular pair has no adjustment. Both tables must not kern it twice.
                if (byScript[_scripts[index]].Count != 0 || _scripts[index] != _scripts[index + 1]) continue;
                int value = _owner.KernPairAdjustment((ushort)_tokens[index].GlyphId,
                    (ushort)_tokens[index + 1].GlyphId, _budget);
                AddPosition(index, 0, value, 0, 0);
            }
        }

        // Attachments and pair positioning currently have separate owners. An
        // adjustment to a mark cannot be reordered across attachment lookups or
        // zeroed without changing glyph geometry. Reject each record before
        // aggregation, so later opposing values cannot conceal that composition.
        private void AddPosition(int index, int x, int advanceX, int y, int advanceY) {
            if (_rejectMarkAdjustments && (x != 0 || y != 0 || advanceX != 0 || advanceY != 0)) {
                OfficeOpenTypeSubstitution.GlyphToken token = _tokens[index];
                if (OfficeManagedTextShaper.IsNonPaintingControl(token.Scalar)
                    || OfficeOpenTypeGlyphFilter.IsUnicodeMark(token.Scalar)
                    || _filter.GetGlyphClass(token.GlyphId, token.Scalar,
                        token.LigatureComponentTextIndexes?.Count > 1, _budget.Charge) == 3)
                    throw new InvalidDataException("GPOS mark adjustment and attachment composition is unsupported.");
            }
            _positions[index] = _positions[index].Add(x, advanceX, y, advanceY);
        }

        private static string[] ResolveScripts(IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> tokens) {
            var result = new string[tokens.Count];
            string previous = "DFLT";
            for (int index = 0; index < tokens.Count; index++) {
                if ((uint)tokens[index].GlyphId > ushort.MaxValue)
                    throw new InvalidDataException("OpenType glyph is out of range.");
                string script = ScriptTagForScalar(tokens[index].Scalar);
                if (script != "DFLT") previous = script;
                result[index] = OfficeOpenTypeGlyphFilter.IsUnicodeMark(tokens[index].Scalar) ? previous : script;
            }
            string following = "DFLT";
            for (int index = tokens.Count - 1; index >= 0; index--) {
                if (result[index] != "DFLT") following = result[index];
                else if (OfficeOpenTypeGlyphFilter.IsUnicodeMark(tokens[index].Scalar)) result[index] = following;
            }
            // Common characters inherit only within one script, never across a mixed-script boundary.
            previous = "DFLT";
            for (int index = 0; index < result.Length; index++) {
                if (result[index] != "DFLT") { previous = result[index]; continue; }
                int end = index;
                while (end + 1 < result.Length && result[end + 1] == "DFLT") end++;
                following = end + 1 < result.Length ? result[end + 1] : "DFLT";
                string inherited = previous == "DFLT" ? following
                    : following == "DFLT" || previous == following ? previous : "DFLT";
                for (; index <= end; index++) result[index] = inherited;
                index--;
            }
            return result;
        }

        private void ValidateSelectionTables() {
            int scripts = Count(_scriptList, 6), features = Count(_featureList, 6);
            Count(_lookupList, 2);
            for (int index = 0; index < scripts; index++) {
                _budget.Charge();
                int script = Relative(_scriptList, U16(_scriptList + 6 + index * 6), 4);
                int languages = U16(script + 2);
                if (languages > MaximumRecords) throw new InvalidDataException("GPOS language limit exceeded.");
                Ensure(script + 4, checked(languages * 6));
                if (U16(script) != 0) ValidateLanguageSystem(Relative(script, U16(script), 6));
                for (int language = 0; language < languages; language++) {
                    _budget.Charge();
                    ValidateLanguageSystem(Relative(script, U16(script + 8 + language * 6), 6));
                }
            }
            for (int index = 0; index < features; index++) {
                _budget.Charge();
                int feature = Relative(_featureList, U16(_featureList + 6 + index * 6), 4);
                int count = U16(feature + 2);
                if (count > MaximumRecords) throw new InvalidDataException("GPOS feature limit exceeded.");
                Ensure(feature + 4, checked(count * 2));
            }
        }

        private void ValidateLanguageSystem(int language) {
            int count = U16(language + 4);
            if (count > MaximumRecords) throw new InvalidDataException("GPOS language feature limit exceeded.");
            Ensure(language + 6, checked(count * 2));
        }

        private int Count(int table, int stride) {
            int count = U16(table);
            if (count > MaximumRecords) throw new InvalidDataException("GPOS record limit exceeded.");
            Ensure(table + 2, checked(count * stride)); return count;
        }
        private int Relative(int table, uint relative, int size) {
            if (relative == 0 || relative > int.MaxValue) throw new InvalidDataException("Invalid GPOS offset.");
            int result = checked(table + (int)relative); Ensure(result, size); return result;
        }
        private void Ensure(int offset, int size) {
            if (!_owner.InPositioningBounds(offset, size, _budget)) throw new InvalidDataException("Truncated GPOS data.");
        }
        private ushort U16(int offset) { Ensure(offset, 2); return _owner.ReadUInt16(offset); }
        private uint U32(int offset) { Ensure(offset, 4); return _owner.ReadUInt32(offset); }
    }
}

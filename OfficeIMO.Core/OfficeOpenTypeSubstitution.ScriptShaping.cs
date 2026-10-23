using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeSubstitution {
    private static readonly string[] ArabicFeatureStages = { "ccmp", "locl", "isol", "fina", "medi", "init", "rlig", "rclt", "calt", "liga", "clig", "mset" };
    private static readonly string[] MarkScriptFeatureStages = { "ccmp", "locl", "rlig", "calt", "liga", "clig" };

    /// <summary>Applies selected Arabic script features to nominal glyphs in logical source order.</summary>
    /// <remarks>
    /// The caller supplies one script segment. Source controls participate in joining even when
    /// omitted from the glyph list. Unsupported lookups decline the entire operation without
    /// publishing a partly shaped run. Visual bidi ordering belongs to the caller afterwards.
    /// </remarks>
    internal bool TryApplyArabicDefaults(List<GlyphToken> glyphs, OfficeTextFeatureSettings settings,
        string? language, CancellationToken cancellationToken, string sourceText, ScriptShapingContext? context = null) =>
        TryApplyScriptDefaults(glyphs, settings, "arab", language, cancellationToken, sourceText, context);

    /// <summary>Composes nominal Latin or Hebrew mark runs through the selected font script before GPOS attachment.</summary>
    internal bool TryApplyMarkScriptDefaults(List<GlyphToken> glyphs, OfficeTextFeatureSettings settings,
        string scriptTag, string? language, CancellationToken cancellationToken, string sourceText, ScriptShapingContext? context = null) {
        if (scriptTag != "latn" && scriptTag != "hebr") return false;
        return TryApplyScriptDefaults(glyphs, settings, scriptTag, language, cancellationToken, sourceText, context);
    }

    private bool TryApplyScriptDefaults(List<GlyphToken> glyphs, OfficeTextFeatureSettings settings,
        string script, string? language, CancellationToken cancellationToken, string sourceText, ScriptShapingContext? context) {
        if (glyphs == null) throw new ArgumentNullException(nameof(glyphs));
        if (settings == null) throw new ArgumentNullException(nameof(settings));
        if (sourceText == null) throw new ArgumentNullException(nameof(sourceText));
        cancellationToken.ThrowIfCancellationRequested();
        try {
            context ??= CreateScriptShapingContext(sourceText, settings, language, cancellationToken);
            string?[]? joining = script == "arab" ? context.Joining : Array.Empty<string?>();
            if (joining == null) return false;
            List<ArabicFeatureLookup>? lookups = context.Lookups(script);
            if (lookups == null) return false;
            var result = new List<GlyphToken>(glyphs);
            context.Charge();
            foreach (ArabicFeatureLookup lookup in lookups) {
                context.Charge();
                ApplyArabicLookup(result, lookup, context.Filter, joining, context,
                    cancellationToken, ref context.Operations);
            }
            glyphs.Clear(); glyphs.AddRange(result);
            return true;
        } catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException
            || exception is ArgumentOutOfRangeException || exception is IndexOutOfRangeException) { return false; }
    }

    private List<ArabicFeatureLookup>? ReadScriptLookups(string scriptTag, string[] stages, ScriptShapingContext context) {
        OfficeTextFeatureSettings settings = context.Settings;
        CancellationToken cancellationToken = context.CancellationToken;
        int[]? features = ReadScriptFeatureIndexes(scriptTag, context.Language, context, out int requiredFeature);
        if (features == null) return null;
        var result = new List<ArabicFeatureLookup>();
        foreach (int featureIndex in features) {
            context.Charge();
            string tag = ReadTag(_featureList + 2 + featureIndex * 6);
            if (scriptTag == "arab" && tag == "stch" && (!settings.TryGetValue(tag, out int stretch) || stretch != 0)) return null;
            if (featureIndex == requiredFeature && Array.IndexOf(stages, tag) < 0) return null;
            if (Array.IndexOf(stages, tag) < 0 && settings.TryGetValue(tag, out int value) && value > 0) return null;
        }
        // Feature order is the Arabic shaping sequence; each feature's lookup order follows
        // LookupList order, with repeated references applied only once in that stage.
        foreach (string stage in stages) {
            context.Charge();
            var stageLookups = new SortedSet<int>();
            int stageValue = settings.TryGetValue(stage, out int authored) ? authored : 1;
            foreach (int featureIndex in features) {
                context.Charge();
                int record = _featureList + 2 + featureIndex * 6;
                if (ReadTag(record) != stage) continue;
                int value = featureIndex == requiredFeature ? 1 : stageValue;
                if (value <= 0) continue;
                int feature = Relative(_featureList, _reader.ReadUInt16(record + 4), 4);
                int count = _reader.ReadUInt16(feature + 2);
                if (count > MaximumLookupRecords) return null;
                Ensure(feature + 4, checked(count * 2));
                for (int index = 0; index < count; index++) {
                    context.Charge();
                    stageLookups.Add(_reader.ReadUInt16(feature + 4 + index * 2));
                }
            }
            bool joining = stage is "isol" or "fina" or "medi" or "init";
            foreach (int lookup in stageLookups) {
                int beforeInspections = context.Inspections;
                if (!CanApplyLookup(lookup, 0, ref context.Inspections, cancellationToken,
                    allowGlyphCountChanges: !joining, glyphFilter: context.Filter)) return null;
                ChargeOperations(ref context.Operations, context.Inspections - beforeInspections);
                result.Add(new ArabicFeatureLookup(lookup, stageValue <= 0 ? 1 : stageValue, joining ? stage : null));
            }
        }
        return result;
    }

    private int[]? ReadScriptFeatureIndexes(string scriptTag, string? language, ScriptShapingContext context, out int requiredFeature) {
        requiredFeature = -1;
        int scriptList = Relative(_table, _reader.ReadUInt16(_table + 4), 2);
        int scriptCount = _reader.ReadUInt16(scriptList);
        if (scriptCount > MaximumFeatureRecords) return null;
        Ensure(scriptList + 2, checked(scriptCount * 6));
        int script = 0;
        for (int index = 0; index < scriptCount; index++) {
            context.Charge();
            int record = scriptList + 2 + index * 6;
            string candidate = ReadTag(record);
            if (candidate == scriptTag) { script = Relative(scriptList, _reader.ReadUInt16(record + 4), 4); break; }
            if (scriptTag != "arab" && candidate == "DFLT") script = Relative(scriptList, _reader.ReadUInt16(record + 4), 4);
        }
        if (script == 0) return scriptTag == "arab" ? null : Array.Empty<int>();
        string? languageTag = ShapingLanguageTag(language);
        int selected = _reader.ReadUInt16(script);
        int languageCount = _reader.ReadUInt16(script + 2);
        if (languageCount > MaximumFeatureRecords) return null;
        Ensure(script + 4, checked(languageCount * 6));
        for (int index = 0; index < languageCount && languageTag != null; index++) {
            context.Charge();
            int record = script + 4 + index * 6;
            if (ReadTag(record) == languageTag) { selected = _reader.ReadUInt16(record + 4); break; }
        }
        if (selected == 0) return null;
        int langSys = Relative(script, selected, 6);
        // Reserved LookupOrder offsets are not an alternate shaping algorithm.
        if (_reader.ReadUInt16(langSys) != 0) return null;
        int count = _reader.ReadUInt16(langSys + 4);
        if (count > MaximumFeatureRecords) return null;
        Ensure(langSys + 6, checked(count * 2));
        Ensure(_featureList, 2);
        int total = _reader.ReadUInt16(_featureList);
        if (total > MaximumFeatureRecords) return null;
        Ensure(_featureList + 2, checked(total * 6));
        var indexes = new SortedSet<int>();
        int required = _reader.ReadUInt16(langSys + 2);
        if (required != ushort.MaxValue) { requiredFeature = required; indexes.Add(required); }
        for (int index = 0; index < count; index++) { context.Charge(); indexes.Add(_reader.ReadUInt16(langSys + 6 + index * 2)); }
        var result = new int[indexes.Count];
        int output = 0;
        foreach (int index in indexes) { context.Charge(); if (index >= total) return null; result[output++] = index; }
        return result;
    }

    private static string? ShapingLanguageTag(string? language) {
        if (string.IsNullOrEmpty(language)) return null;
        string primary = language!.Split('-')[0].ToLowerInvariant();
        return primary switch {
            "ar" => "ARA ", "fa" => "FAR ", "ur" => "URD ", "ps" => "PAS ",
            "sd" => "SND ", "ug" => "UYG ", "ku" => "KUR ", "he" => "IWR ",
            "yi" => "JII ", "tr" => "TRK ", "az" => "AZE ", "lt" => "LTH ",
            "nl" => "NLD ", "de" => "DEU ", "en" => "ENG ", _ => null
        };
    }

    private void ApplyArabicLookup(List<GlyphToken> glyphs, ArabicFeatureLookup feature,
        OfficeOpenTypeGlyphFilter filter, string?[] joining, ScriptShapingContext context,
        CancellationToken cancellationToken, ref int operations) {
        // Normalized glyphs can share an original grapheme's TextIndex. Split on
        // the retained scalar before filtering so ignored marks cannot migrate
        // across its barrier when they are restored into the substituted run.
        bool controls = false;
        foreach (GlyphToken token in glyphs) {
            cancellationToken.ThrowIfCancellationRequested();
            ChargeOperations(ref operations, 1);
            if (IsJoiningControl(token.Scalar)) { controls = true; break; }
        }
        if (!controls) {
            ApplyArabicLookupWithoutControls(glyphs, feature, filter, joining, context,
                cancellationToken, ref operations);
            return;
        }
        var result = new List<GlyphToken>(glyphs.Count);
        int start = 0;
        for (int index = 0; index <= glyphs.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            ChargeOperations(ref operations, 1);
            if (index < glyphs.Count && !IsJoiningControl(glyphs[index].Scalar)) continue;
            if (index > start) {
                List<GlyphToken> segment = glyphs.GetRange(start, index - start);
                ApplyArabicLookupWithoutControls(segment, feature, filter, joining, context,
                    cancellationToken, ref operations);
                result.AddRange(segment);
            }
            // The caller has already admitted this non-painting glyph. Neither
            // substitutions nor GDEF filtering may drop its source ownership.
            if (index < glyphs.Count) result.Add(glyphs[index]);
            start = index + 1;
        }
        glyphs.Clear();
        glyphs.AddRange(result);
    }

    private static bool IsJoiningControl(int scalar) => scalar is 0x200C or 0x200D;

    private void ApplyArabicLookupWithoutControls(List<GlyphToken> glyphs, ArabicFeatureLookup feature,
        OfficeOpenTypeGlyphFilter filter, string?[] joining, ScriptShapingContext context,
        CancellationToken cancellationToken, ref int operations) {
        int lookup = Relative(_lookupList, _reader.ReadUInt16(_lookupList + 2 + feature.LookupIndex * 2), 6);
        ushort flags = _reader.ReadUInt16(lookup + 2);
        int subtableCount = _reader.ReadUInt16(lookup + 4);
        int set = (flags & 0x10) != 0 ? _reader.ReadUInt16(lookup + 6 + subtableCount * 2) : -1;
        var included = new List<GlyphToken>(glyphs.Count);
        var ignored = new List<GlyphToken>();
        // IsIgnored calls this bounded closure for every GDEF/filter-table inspection.
        int filterOperations = operations;
        Action charge = () => { cancellationToken.ThrowIfCancellationRequested(); ChargeOperations(ref filterOperations, 1); };
        object lookupIdentity = new object();
        int lookupOrder = 0;
        foreach (GlyphToken original in glyphs) {
            cancellationToken.ThrowIfCancellationRequested();
            GlyphToken token = original.WithLookupOrder(lookupOrder++, lookupIdentity);
            if (filter.IsIgnored(token, flags, set, charge)) ignored.Add(token);
            else included.Add(token);
        }
        operations = filterOperations;
        if (!context.HasJoiningControls) {
            ApplyArabicLookupSegment(included, feature, filter, joining, cancellationToken, ref operations);
            RestoreArabicIgnoredGlyphs(glyphs, included, ignored, context.Source, cancellationToken, ref operations);
            return;
        }
        var substituted = new List<GlyphToken>(included.Count);
        int segmentStart = 0;
        for (int boundary = 1; boundary <= included.Count; boundary++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (boundary < included.Count && !context.HasJoiningControl(included[boundary - 1], included[boundary])) continue;
            var segment = included.GetRange(segmentStart, boundary - segmentStart);
            ApplyArabicLookupSegment(segment, feature, filter, joining, cancellationToken, ref operations);
            substituted.AddRange(segment);
            segmentStart = boundary;
        }
        RestoreArabicIgnoredGlyphs(glyphs, substituted, ignored, context.Source, cancellationToken, ref operations);
    }

    private void ApplyArabicLookupSegment(List<GlyphToken> segment, ArabicFeatureLookup feature,
        OfficeOpenTypeGlyphFilter filter, string?[] joining,
        CancellationToken cancellationToken, ref int operations) {
        if (feature.JoiningForm == null) {
            ApplyLookup(segment, feature.LookupIndex, feature.Value, cancellationToken, ref operations);
            return;
        }
        int filterOperations = operations;
        Action charge = () => { cancellationToken.ThrowIfCancellationRequested(); ChargeOperations(ref filterOperations, 1); };
        for (int index = 0; index < segment.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            GlyphToken token = segment[index];
            // A normalized mark shares its original base's source index. It
            // must not inherit that letter's init/medi/fina substitution mask.
            if (OfficeOpenTypeGlyphFilter.IsUnicodeMark(token.Scalar)) continue;
            filterOperations = operations;
            bool mark = filter.GetGlyphClass(token.GlyphId, token.Scalar,
                token.LigatureComponentTextIndexes?.Count > 1, charge) == 3;
            operations = filterOperations;
            if (mark) continue;
            int sourceIndex = token.TextIndex;
            if ((uint)sourceIndex < (uint)joining.Length && joining[sourceIndex] == feature.JoiningForm)
                ApplyLookupAt(segment, index, feature.LookupIndex, feature.Value, cancellationToken, ref operations, 0);
        }
    }

    private static void RestoreArabicIgnoredGlyphs(List<GlyphToken> destination, List<GlyphToken> included,
        List<GlyphToken> ignored, string sourceText, CancellationToken cancellationToken, ref int operations) {
        // A ligature skipping a mark owns a contiguous logical source span. Its retained mark
        // glyph remains a continuation with the original source index for component attachment.
        for (int index = 0; index < included.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            GlyphToken owner = included[index];
            if (owner.LigatureComponentTextIndexes == null) continue;
            int end = owner.TextIndex + owner.UnicodeText.Length;
            foreach (int component in owner.LigatureComponentTextIndexes) {
                ChargeOperations(ref operations, 1);
                if ((uint)component >= (uint)sourceText.Length) throw new InvalidDataException("A GSUB component has no logical source.");
                end = Math.Max(end, component + (char.IsHighSurrogate(sourceText[component])
                    && component + 1 < sourceText.Length && char.IsLowSurrogate(sourceText[component + 1]) ? 2 : 1));
            }
            if (end > sourceText.Length) throw new InvalidDataException("A GSUB ligature exceeds its logical source.");
            // A mark-only ccmp output can be an empty continuation at its base's
            // source index. Rebuilding it from that index would invent ownership.
            if (!owner.IsUnicodeContinuation) {
                included[index] = new GlyphToken(owner.GlyphId, sourceText.Substring(owner.TextIndex, end - owner.TextIndex),
                    owner.TextIndex, owner.Scalar, false, owner.ClusterStart, owner.LigatureComponentTextIndexes,
                    owner.LigatureIdentity, owner.LigatureComponent, owner.LookupOrder, owner.LookupIdentity);
            }
            for (int markIndex = 0; markIndex < ignored.Count; markIndex++) {
                ChargeOperations(ref operations, 1);
                GlyphToken mark = ignored[markIndex];
                int component = 0;
                bool associated = owner.LigatureIdentity is LigatureAssociation association
                    && association.TryAssociate(mark, ref operations, out component);
                bool covered = !owner.IsUnicodeContinuation && mark.TextIndex > owner.TextIndex && mark.TextIndex < end;
                if (!associated && !covered) continue;
                ignored[markIndex] = new GlyphToken(mark.GlyphId, covered ? string.Empty : mark.UnicodeText,
                    mark.TextIndex, mark.Scalar, covered || mark.IsUnicodeContinuation,
                    associated || covered ? owner.ClusterStart : mark.ClusterStart, mark.LigatureComponentTextIndexes,
                    associated ? owner.LigatureIdentity : mark.LigatureIdentity,
                    associated ? component : mark.LigatureComponent, mark.LookupOrder, mark.LookupIdentity);
            }
        }
        destination.Clear();
        int ignoredIndex = 0;
        foreach (GlyphToken token in included) {
            ChargeOperations(ref operations, 1);
            while (ignoredIndex < ignored.Count && ignored[ignoredIndex].LookupOrder < token.LookupOrder)
                destination.Add(ignored[ignoredIndex++]);
            destination.Add(token);
        }
        while (ignoredIndex < ignored.Count) destination.Add(ignored[ignoredIndex++]);
    }

    internal readonly struct ArabicFeatureLookup {
        internal ArabicFeatureLookup(int lookupIndex, int value, string? joiningForm) {
            LookupIndex = lookupIndex;
            Value = value;
            JoiningForm = joiningForm;
        }
        internal int LookupIndex { get; }
        internal int Value { get; }
        internal string? JoiningForm { get; }
    }
}

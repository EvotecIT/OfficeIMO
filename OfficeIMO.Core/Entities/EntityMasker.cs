using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Masks findings using their original source coordinates.</summary>
public static class EntityMasker {
    /// <summary>Replaces selected findings in one source segment using the default masking options.</summary>
    public static string Mask(string text, IEnumerable<EntityMatch> matches) => Mask(text, matches, options: null);

    /// <summary>Replaces selected findings contained in one source segment.</summary>
    public static string Mask(string text, IEnumerable<EntityMatch> matches, EntityMaskOptions? options) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        return Replace(new EntityMaskingSource(text), matches, options ?? new EntityMaskOptions(), forcePseudonym: false, map: null)[0];
    }

    /// <summary>Replaces selected findings across source segments using the default masking options.</summary>
    public static IReadOnlyList<string> Mask(IReadOnlyList<EntityTextSegment> segments, IEnumerable<EntityMatch> matches) =>
        Mask(segments, matches, options: null);

    /// <summary>Replaces selected findings across source segments and returns the text for each segment in order.</summary>
    /// <remarks>Segmented fixed masking replaces each matched character with one mask character.</remarks>
    public static IReadOnlyList<string> Mask(IReadOnlyList<EntityTextSegment> segments, IEnumerable<EntityMatch> matches,
        EntityMaskOptions? options) {
        if (segments == null) throw new ArgumentNullException(nameof(segments));
        return Replace(new EntityMaskingSource(segments), matches, options ?? new EntityMaskOptions(), forcePseudonym: false, map: null);
    }

    /// <summary>Replaces selected findings with keyed pseudonyms and returns the sensitive token map.</summary>
    public static EntityPseudonymResult Pseudonymize(string text, IEnumerable<EntityMatch> matches, EntityMaskOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (text == null) throw new ArgumentNullException(nameof(text));
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        string masked = Replace(new EntityMaskingSource(text), matches, options, forcePseudonym: true, map)[0];
        return new EntityPseudonymResult(masked, map);
    }

    private static IReadOnlyList<string> Replace(EntityMaskingSource source, IEnumerable<EntityMatch> matches, EntityMaskOptions options,
        bool forcePseudonym, Dictionary<string, string>? map) {
        if (matches == null) throw new ArgumentNullException(nameof(matches));
        if (options.FixedLength < 0) throw new ArgumentOutOfRangeException(nameof(options.FixedLength));

        EntityMaskStyle style = forcePseudonym ? EntityMaskStyle.Pseudonym : options.Style;
        if (style is not (EntityMaskStyle.KindLabel or EntityMaskStyle.Fixed or EntityMaskStyle.Partial or EntityMaskStyle.Pseudonym))
            throw new ArgumentOutOfRangeException(nameof(options.Style));
        byte[]? key = options.PseudonymKey;
        if (style == EntityMaskStyle.Pseudonym && (key == null || key.Length == 0))
            throw new ArgumentException("A non-empty key is required to create pseudonyms.", nameof(options));
        using HMACSHA256? hmac = style == EntityMaskStyle.Pseudonym ? new HMACSHA256(key!) : null;

        var ordered = new List<Reading>();
        foreach (EntityMatch match in matches) AddSelected(source, match, ordered, options);
        ordered.Sort(CompareForMasking);

        for (int index = 0; index < ordered.Count;) {
            int startIndex = index;
            Reading first = ordered[index++];
            long end = first.End;
            bool aggregate = false;
            while (index < ordered.Count && ordered[index].Chain == first.Chain && ordered[index].Start <= end) {
                Reading next = ordered[index++];
                aggregate |= next.Start != first.Start || next.End != first.End || !ReferenceEquals(first.Match.Scheme, next.Match.Scheme);
                end = Math.Max(end, next.End);
            }

            string? replacement = Replacement(source, ordered, startIndex, index, end, aggregate, style, options, hmac, map);
            source.Apply(first.Start, end, replacement, style == EntityMaskStyle.Partial && !aggregate, options.FixedCharacter);
        }

        return source.GetResult();
    }

    private static void AddSelected(EntityMaskingSource source, EntityMatch match, List<Reading> ordered, EntityMaskOptions options) {
        var range = source.GetRange(match);
        bool selected = options.Sensitivities.Contains(match.Scheme.Sensitivity);
        if (selected) ordered.Add(new Reading(match, range));
        if (match.MaskingFallbacks.Count == 0 && (selected || match.Contained.Count == 0)) return;

        var pending = new Stack<EntityMatch>();
        PushFallbacks(match, pending, selected);
        while (pending.Count > 0) {
            EntityMatch next = pending.Pop();
            bool include = options.Sensitivities.Contains(next.Scheme.Sensitivity);
            if (include) ordered.Add(new Reading(next, source.GetRange(next)));
            PushFallbacks(next, pending, include);
        }
    }

    private static void PushFallbacks(EntityMatch match, Stack<EntityMatch> pending, bool selected) {
        if (!selected) {
            foreach (EntityMatch item in match.Contained) pending.Push(item);
        }
        foreach (EntityMatch item in match.MaskingFallbacks) {
            if (!selected || !Covers(match, item)) pending.Push(item);
        }
    }

    private static bool Covers(EntityMatch outer, EntityMatch inner) {
        EntityMatchPart first = outer.Parts[0];
        EntityMatchPart last = outer.Parts[outer.Parts.Count - 1];
        EntityMatchPart innerFirst = inner.Parts[0];
        EntityMatchPart innerLast = inner.Parts[inner.Parts.Count - 1];
        return (first.SegmentIndex, first.Start).CompareTo((innerFirst.SegmentIndex, innerFirst.Start)) <= 0 &&
               (last.SegmentIndex, last.Start + last.Length).CompareTo((innerLast.SegmentIndex, innerLast.Start + innerLast.Length)) >= 0;
    }

    private static string? Replacement(EntityMaskingSource source, List<Reading> ordered, int startIndex, int endIndex, long end,
        bool aggregate, EntityMaskStyle style, EntityMaskOptions options, HMACSHA256? hmac, Dictionary<string, string>? map) {
        Reading first = ordered[startIndex];
        if (style == EntityMaskStyle.KindLabel) return "[" + first.Match.Kind + "]";
        if (style == EntityMaskStyle.Partial && !aggregate) return first.Match.Scheme.MaskPartial(first.Match.Normalized);
        if (style is EntityMaskStyle.Fixed or EntityMaskStyle.Partial) {
            // A null replacement fills each covered source slice at its written length.
            return source.IsSingleText && options.FixedLength is int length ? new string(options.FixedCharacter, length) : null;
        }

        string keyMaterial = aggregate ? AggregateKey(ordered, startIndex, endIndex) : first.Match.Normalized;
        byte[] encoded = Encoding.UTF8.GetBytes(keyMaterial);
        if (aggregate) {
            // Distinguish a combined payload from a single canonical value with a non-UTF-8 byte.
            Array.Resize(ref encoded, encoded.Length + 1);
            encoded[encoded.Length - 1] = 0xFF;
        }
        byte[] hash = hmac!.ComputeHash(encoded);
        var token = new StringBuilder().Append('[').Append(aggregate ? "Entities" : first.Match.Kind.ToString()).Append(':');
        const string hex = "0123456789abcdef";
        for (int index = 0; index < 16; index++) {
            token.Append(hex[hash[index] >> 4]).Append(hex[hash[index] & 15]);
        }
        string pseudonym = token.Append(']').ToString();
        if (map != null && !map.ContainsKey(pseudonym))
            map.Add(pseudonym, aggregate ? source.OriginalSingleText.Substring((int)first.Start, (int)(end - first.Start)) : first.Match.Normalized);
        return pseudonym;
    }

    private static string AggregateKey(List<Reading> ordered, int start, int end) {
        var readings = new SortedSet<string>(StringComparer.Ordinal);
        for (int index = start; index < end; index++) {
            EntityMatch match = ordered[index].Match;
            readings.Add(match.Scheme.Id + "\0" + match.Normalized);
        }
        var key = new StringBuilder();
        foreach (string reading in readings) {
            key.Append(reading.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(reading);
        }
        return key.ToString();
    }

    private static int CompareForMasking(Reading left, Reading right) {
        int start = left.Start.CompareTo(right.Start);
        if (start != 0) return start;
        int length = right.End.CompareTo(left.End);
        if (length != 0) return length;
        int confidence = right.Match.Confidence.CompareTo(left.Match.Confidence);
        return confidence != 0 ? confidence : StringComparer.OrdinalIgnoreCase.Compare(left.Match.Scheme.Id, right.Match.Scheme.Id);
    }

    private readonly struct Reading {
        internal Reading(EntityMatch match, (long Start, long End, int Chain) range) {
            Match = match;
            Start = range.Start;
            End = range.End;
            Chain = range.Chain;
        }

        internal EntityMatch Match { get; }
        internal long Start { get; }
        internal long End { get; }
        internal int Chain { get; }
    }
}

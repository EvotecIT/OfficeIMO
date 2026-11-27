using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Entities;

/// <summary>Masks findings for a single source segment using their original offsets.</summary>
public static class EntityMasker {
    /// <summary>Replaces selected findings with the requested masking style.</summary>
    public static string Mask(string text, IEnumerable<EntityMatch> matches, EntityMaskOptions? options = null) =>
        Replace(text, matches, options ?? new EntityMaskOptions(), forcePseudonym: false, map: null);

    /// <summary>Replaces selected findings with keyed pseudonyms and returns the sensitive token map.</summary>
    public static EntityPseudonymResult Pseudonymize(string text, IEnumerable<EntityMatch> matches, EntityMaskOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        return new EntityPseudonymResult(Replace(text, matches, options, forcePseudonym: true, map: map), map);
    }

    private static string Replace(string text, IEnumerable<EntityMatch> matches, EntityMaskOptions options,
        bool forcePseudonym, Dictionary<string, string>? map) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (matches == null) throw new ArgumentNullException(nameof(matches));
        if (options.FixedLength < 0) throw new ArgumentOutOfRangeException(nameof(options.FixedLength));

        EntityMaskStyle style = forcePseudonym ? EntityMaskStyle.Pseudonym : options.Style;
        if (style is not (EntityMaskStyle.KindLabel or EntityMaskStyle.Fixed or EntityMaskStyle.Partial or EntityMaskStyle.Pseudonym))
            throw new ArgumentOutOfRangeException(nameof(options.Style));
        byte[]? key = options.PseudonymKey;
        if (style == EntityMaskStyle.Pseudonym && (key == null || key.Length == 0))
            throw new ArgumentException("A non-empty key is required to create pseudonyms.", nameof(options));
        using HMACSHA256? hmac = style == EntityMaskStyle.Pseudonym ? new HMACSHA256(key!) : null;

        var ordered = new List<EntityMatch>();
        int? segment = null;
        foreach (EntityMatch match in matches) {
            if (match == null || match.Start > text.Length - match.Length)
                throw new ArgumentException("A match falls outside the supplied text.", nameof(matches));
            if (segment.HasValue && segment.Value != match.SegmentIndex)
                throw new ArgumentException("Supply matches from one text segment per masking operation.", nameof(matches));
            segment = match.SegmentIndex;
            AddSelected(match, ordered, options);
        }
        if (ordered.Count == 0) return text;

        ordered.Sort(CompareForMasking);
        var output = new StringBuilder(text.Length);
        int copied = 0;
        for (int index = 0; index < ordered.Count;) {
            int startIndex = index;
            EntityMatch first = ordered[index++];
            int end = first.Start + first.Length;
            bool aggregate = false;
            while (index < ordered.Count && ordered[index].Start <= end) {
                EntityMatch next = ordered[index++];
                aggregate |= next.Start != first.Start || next.Length != first.Length || !ReferenceEquals(first.Scheme, next.Scheme);
                end = Math.Max(end, next.Start + next.Length);
            }
            output.Append(text, copied, first.Start - copied);
            output.Append(Replacement(text, ordered, startIndex, index, end, aggregate, style, options, hmac, map));
            copied = end;
        }

        output.Append(text, copied, text.Length - copied);
        return output.ToString();
    }

    // A skipped winner can still cover a sensitive reading such as the mailbox in a mailto URL.
    private static void AddSelected(EntityMatch match, List<EntityMatch> ordered, EntityMaskOptions options) {
        if (options.Sensitivities.Contains(match.Scheme.Sensitivity)) {
            ordered.Add(match);
            return;
        }
        if (match.Contained.Count == 0 && match.MaskingFallbacks.Count == 0) return;
        var pending = new Stack<EntityMatch>();
        PushFallbacks(match, pending);
        while (pending.Count > 0) {
            EntityMatch next = pending.Pop();
            if (options.Sensitivities.Contains(next.Scheme.Sensitivity)) ordered.Add(next);
            else PushFallbacks(next, pending);
        }
    }

    private static void PushFallbacks(EntityMatch match, Stack<EntityMatch> pending) {
        foreach (EntityMatch item in match.Contained) pending.Push(item);
        foreach (EntityMatch item in match.MaskingFallbacks) pending.Push(item);
    }

    private static string Replacement(string text, List<EntityMatch> ordered, int startIndex, int endIndex, int end,
        bool aggregate, EntityMaskStyle style, EntityMaskOptions options, HMACSHA256? hmac, Dictionary<string, string>? map) {
        EntityMatch first = ordered[startIndex];
        int length = end - first.Start;
        if (style == EntityMaskStyle.KindLabel) return "[" + first.Kind + "]";
        if (style == EntityMaskStyle.Fixed) return new string(options.FixedCharacter, options.FixedLength ?? length);
        if (style == EntityMaskStyle.Partial) return aggregate
            ? new string(options.FixedCharacter, options.FixedLength ?? length)
            : first.Scheme.MaskPartial(first.Normalized);

        string keyMaterial = aggregate ? AggregateKey(ordered, startIndex, endIndex) : first.Normalized;
        byte[] encoded = Encoding.UTF8.GetBytes(keyMaterial);
        if (aggregate) {
            // Distinguish a combined payload from a single canonical value with a non-UTF-8 byte.
            Array.Resize(ref encoded, encoded.Length + 1);
            encoded[encoded.Length - 1] = 0xFF;
        }
        byte[] hash = hmac!.ComputeHash(encoded);
        var token = new StringBuilder().Append('[').Append(aggregate ? "Entities" : first.Kind.ToString()).Append(':');
        const string hex = "0123456789abcdef";
        for (int index = 0; index < 16; index++) {
            token.Append(hex[hash[index] >> 4]).Append(hex[hash[index] & 15]);
        }
        string pseudonym = token.Append(']').ToString();
        if (map != null && !map.ContainsKey(pseudonym))
            map.Add(pseudonym, aggregate ? text.Substring(first.Start, length) : first.Normalized);
        return pseudonym;
    }

    private static string AggregateKey(List<EntityMatch> ordered, int start, int end) {
        var readings = new SortedSet<string>(StringComparer.Ordinal);
        for (int index = start; index < end; index++) {
            EntityMatch reading = ordered[index];
            readings.Add(reading.Scheme.Id + "\0" + reading.Normalized);
        }
        var key = new StringBuilder();
        foreach (string reading in readings) {
            key.Append(reading.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(':').Append(reading);
        }
        return key.ToString();
    }

    private static int CompareForMasking(EntityMatch left, EntityMatch right) {
        int start = left.Start.CompareTo(right.Start);
        if (start != 0) return start;
        int length = right.Length.CompareTo(left.Length);
        if (length != 0) return length;
        int confidence = right.Confidence.CompareTo(left.Confidence);
        return confidence != 0 ? confidence : StringComparer.OrdinalIgnoreCase.Compare(left.Scheme.Id, right.Scheme.Id);
    }
}

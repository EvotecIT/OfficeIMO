using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>Immutable bounded set of Unicode scalar ranges used by scoped font selection.</summary>
public sealed class OfficeFontUnicodeRangeSet : IEquatable<OfficeFontUnicodeRangeSet> {
    private const int MaximumRangeCount = 128;
    private readonly OfficeFontUnicodeRange[] _ranges;
    private readonly ReadOnlyCollection<OfficeFontUnicodeRange> _view;

    /// <summary>All Unicode scalar values.</summary>
    public static OfficeFontUnicodeRangeSet All { get; } =
        new OfficeFontUnicodeRangeSet(new[] { new OfficeFontUnicodeRange(0, 0x10FFFF) });

    /// <summary>Creates a normalized range set.</summary>
    public OfficeFontUnicodeRangeSet(IEnumerable<OfficeFontUnicodeRange> ranges) {
        if (ranges == null) throw new ArgumentNullException(nameof(ranges));
        var ordered = new List<OfficeFontUnicodeRange>(ranges);
        if (ordered.Count == 0 || ordered.Count > MaximumRangeCount) {
            throw new ArgumentOutOfRangeException(nameof(ranges), $"A font range set must contain between 1 and {MaximumRangeCount} ranges.");
        }
        ordered.Sort((left, right) => left.Start != right.Start
            ? left.Start.CompareTo(right.Start)
            : left.End.CompareTo(right.End));
        var merged = new List<OfficeFontUnicodeRange>(ordered.Count);
        foreach (OfficeFontUnicodeRange range in ordered) {
            if (merged.Count == 0 || range.Start > merged[merged.Count - 1].End + 1) {
                merged.Add(range);
            } else {
                OfficeFontUnicodeRange previous = merged[merged.Count - 1];
                merged[merged.Count - 1] = new OfficeFontUnicodeRange(previous.Start, Math.Max(previous.End, range.End));
            }
        }
        _ranges = merged.ToArray();
        _view = new ReadOnlyCollection<OfficeFontUnicodeRange>(_ranges);
    }

    /// <summary>Normalized, ordered ranges.</summary>
    public IReadOnlyList<OfficeFontUnicodeRange> Ranges => _view;

    /// <summary>True when this set covers every Unicode scalar.</summary>
    public bool IsAll => _ranges.Length == 1 && _ranges[0].Start == 0 && _ranges[0].End == 0x10FFFF;

    /// <summary>Returns true when the scalar is included.</summary>
    public bool Contains(int scalar) {
        for (int index = 0; index < _ranges.Length; index++) {
            if (scalar < _ranges[index].Start) return false;
            if (_ranges[index].Contains(scalar)) return true;
        }
        return false;
    }

    /// <summary>Returns true when every scalar in the supplied text is included.</summary>
    public bool ContainsText(string? text) => ContainsTextCore(text, ignoreFontCoverageControls: false);

    internal bool ContainsFontCoverageText(string? text) =>
        ContainsTextCore(text, ignoreFontCoverageControls: true);

    private bool ContainsTextCore(string? text, bool ignoreFontCoverageControls) {
        if (string.IsNullOrEmpty(text)) return true;
        for (int index = 0; index < text!.Length; index++) {
            int scalar = text[index];
            if (char.IsHighSurrogate(text[index])
                && index + 1 < text.Length
                && char.IsLowSurrogate(text[index + 1])) {
                scalar = char.ConvertToUtf32(text[index], text[++index]);
            }
            if (ignoreFontCoverageControls && OfficeTextElements.IsIgnorableFontCoverageScalar(scalar)) continue;
            if (!Contains(scalar)) return false;
        }
        return true;
    }

    /// <summary>Parses a CSS unicode-range descriptor such as U+0000-00FF or U+4??.</summary>
    public static bool TryParseCss(string? value, out OfficeFontUnicodeRangeSet? ranges) {
        ranges = null;
        if (string.IsNullOrWhiteSpace(value) || value!.Length > 4096) return false;
        string[] parts = value.Split(',');
        if (parts.Length == 0 || parts.Length > MaximumRangeCount) return false;
        var parsed = new List<OfficeFontUnicodeRange>(parts.Length);
        foreach (string raw in parts) {
            string token = raw.Trim();
            if (!token.StartsWith("u+", StringComparison.OrdinalIgnoreCase)) return false;
            string body = token.Substring(2);
            int wildcard = body.IndexOf('?');
            if (wildcard >= 0) {
                if (body.Length == 0 || body.Length > 6) return false;
                for (int index = wildcard; index < body.Length; index++) {
                    if (body[index] != '?') return false;
                }
                string prefix = body.Substring(0, wildcard);
                if (!TryHex(prefix + new string('0', body.Length - wildcard), out int start)
                    || !TryHex(prefix + new string('F', body.Length - wildcard), out int end)
                    || end > 0x10FFFF) return false;
                parsed.Add(new OfficeFontUnicodeRange(start, end));
                continue;
            }
            string[] bounds = body.Split('-');
            if (bounds.Length < 1 || bounds.Length > 2
                || !TryHex(bounds[0], out int first)
                || !TryHex(bounds.Length == 2 ? bounds[1] : bounds[0], out int last)
                || first > last || last > 0x10FFFF) return false;
            parsed.Add(new OfficeFontUnicodeRange(first, last));
        }
        ranges = new OfficeFontUnicodeRangeSet(parsed);
        return true;
    }

    internal string ToStableKey() {
        var parts = new string[_ranges.Length];
        for (int index = 0; index < _ranges.Length; index++) {
            parts[index] = _ranges[index].Start.ToString("X", CultureInfo.InvariantCulture)
                + "-"
                + _ranges[index].End.ToString("X", CultureInfo.InvariantCulture);
        }
        return string.Join(",", parts);
    }

    /// <inheritdoc />
    public bool Equals(OfficeFontUnicodeRangeSet? other) =>
        other != null && string.Equals(ToStableKey(), other.ToStableKey(), StringComparison.Ordinal);

    /// <inheritdoc />
    public override bool Equals(object? obj) => Equals(obj as OfficeFontUnicodeRangeSet);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            for (int index = 0; index < _ranges.Length; index++) {
                hash = (hash * 31) + _ranges[index].GetHashCode();
            }
            return hash;
        }
    }

    private static bool TryHex(string value, out int scalar) {
        scalar = 0;
        return value.Length >= 1
            && value.Length <= 6
            && int.TryParse(value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out scalar);
    }
}

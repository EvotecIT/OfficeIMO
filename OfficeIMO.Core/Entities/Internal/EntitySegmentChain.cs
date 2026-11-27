using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Entities.Internal;

/// <summary>Maps a logical search string back to the non-empty segments that supplied its characters.</summary>
internal sealed class EntitySegmentChain {
    private readonly List<(EntityTextSegment Segment, int Index, int End)> _segments = new();
    private string? _first;
    private StringBuilder? _joined;

    internal EntitySegmentChain(EntityTextSegment first, int firstIndex) {
        First = first;
        FirstIndex = firstIndex;
    }

    internal EntityTextSegment First { get; }
    internal int FirstIndex { get; }
    internal int Length { get; private set; }
    internal string GetText() => _joined?.ToString() ?? _first ?? string.Empty;

    internal void Append(EntityTextSegment segment, int index, int length) {
        if (length == 0) return;
        string text = segment.Text;
        if (_first == null) {
            _first = length == text.Length ? text : text.Substring(0, length);
        } else {
            if (_joined == null) _joined = new StringBuilder(_first);
            _joined.Append(text, 0, length);
        }

        Length += length;
        _segments.Add((segment, index, Length));
    }

    internal EntityTextSegment SegmentAt(int offset) =>
        offset < Length ? _segments[IndexAt(offset)].Segment : First;

    internal bool CoversStructuredSegment(int start, int end) {
        int index = IndexAt(start);
        var entry = _segments[index];
        int offset = index == 0 ? 0 : _segments[index - 1].End;
        return entry.Segment.StructuredSource && start == offset && end == entry.End &&
               end - start == entry.Segment.Text.Length;
    }

    internal EntityMatchPart[] GetParts(int start, int end) {
        int first = IndexAt(start);
        int last = IndexAt(end - 1);
        var parts = new EntityMatchPart[last - first + 1];
        for (int index = first; index <= last; index++) {
            var entry = _segments[index];
            int offset = index == 0 ? 0 : _segments[index - 1].End;
            int from = Math.Max(start, offset);
            int to = Math.Min(end, entry.End);
            parts[index - first] = new EntityMatchPart(entry.Index, from - offset, to - from);
        }
        return parts;
    }

    private int IndexAt(int offset) {
        int low = 0;
        int high = _segments.Count - 1;
        while (low < high) {
            int middle = low + (high - low) / 2;
            if (_segments[middle].End <= offset) low = middle + 1;
            else high = middle;
        }
        return low;
    }
}

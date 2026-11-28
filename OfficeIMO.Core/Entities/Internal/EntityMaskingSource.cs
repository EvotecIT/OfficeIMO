using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Entities.Internal;

/// <summary>Projects mask intervals onto the original segments without joining their text.</summary>
internal sealed class EntityMaskingSource {
    private readonly string[] _texts;
    private readonly long[] _starts;
    private readonly int[] _chains;
    private readonly StringBuilder?[] _outputs;
    private readonly int[] _copied;
    private int? _singleSegmentIndex;

    internal EntityMaskingSource(string text) : this(1) {
        _texts[0] = text;
        IsSingleText = true;
    }

    internal EntityMaskingSource(IReadOnlyList<EntityTextSegment> segments) : this(segments.Count) {
        long offset = 0;
        int chain = 0;
        for (int index = 0; index < segments.Count; index++) {
            EntityTextSegment segment = segments[index];
            if (!segment.ContinuesPrevious) chain = index;
            _texts[index] = segment.Text ?? string.Empty;
            _starts[index] = offset;
            _chains[index] = chain;
            offset += _texts[index].Length;
        }
    }

    private EntityMaskingSource(int count) {
        _texts = new string[count];
        _starts = new long[count];
        _chains = new int[count];
        _outputs = new StringBuilder?[count];
        _copied = new int[count];
    }

    internal bool IsSingleText { get; }
    internal string OriginalSingleText => _texts[0];

    internal (long Start, long End, int Chain) GetRange(EntityMatch match) {
        if (match == null) throw new ArgumentException("A match is missing.", "matches");

        IReadOnlyList<EntityMatchPart> parts = match.Parts;
        if (IsSingleText) {
            int segmentIndex = parts[0].SegmentIndex;
            if (parts.Count != 1 || _singleSegmentIndex.HasValue && _singleSegmentIndex.Value != segmentIndex)
                throw new ArgumentException("Supply matches from one text segment per masking operation.", "matches");
            _singleSegmentIndex = segmentIndex;
        }

        long start = 0;
        long end = 0;
        int chain = 0;
        int previous = -1;
        foreach (EntityMatchPart part in parts) {
            int index = IsSingleText ? 0 : part.SegmentIndex;
            if (index <= previous || index >= _texts.Length || part.Start > _texts[index].Length - part.Length)
                throw new ArgumentException("A match part falls outside the supplied text segments.", "matches");

            long partStart = _starts[index] + part.Start;
            if (previous < 0) {
                start = partStart;
                chain = _chains[index];
            } else if (end != partStart || _chains[index] != chain) {
                throw new ArgumentException("A match must cover consecutive text in one segment chain.", "matches");
            }

            previous = index;
            end = partStart + part.Length;
        }
        return (start, end, chain);
    }

    internal void Apply(long start, long end, string? replacement, bool distribute, char fixedCharacter) {
        int first = SegmentAt(start);
        int last = SegmentAt(end - 1);
        long length = end - start;
        int suffix = distribute && first != last && replacement != null && replacement.Length != length
            ? SharedSuffix(start, last, end, replacement)
            : 0;
        int written = 0;
        for (int index = first; index <= last; index++) {
            string text = _texts[index];
            int from = index == first ? (int)(start - _starts[index]) : 0;
            int to = index == last ? (int)(end - _starts[index]) : text.Length;
            if (from == to) continue;

            StringBuilder output = _outputs[index] ??= new StringBuilder(text.Length);
            output.Append(text, _copied[index], from - _copied[index]);
            if (replacement == null) {
                output.Append(fixedCharacter, to - from);
            } else if (distribute) {
                long consumed = _starts[index] + to - start;
                int cut = PartialCut(consumed, length, replacement.Length, suffix);
                output.Append(replacement, written, cut - written);
                written = cut;
            } else if (index == first) {
                output.Append(replacement);
            }
            _copied[index] = to;
        }
    }

    internal IReadOnlyList<string> GetResult() {
        var result = new string[_texts.Length];
        for (int index = 0; index < result.Length; index++) {
            string text = _texts[index];
            StringBuilder? output = _outputs[index];
            result[index] = output == null ? text : output.Append(text, _copied[index], text.Length - _copied[index]).ToString();
        }
        return Array.AsReadOnly(result);
    }

    private int SegmentAt(long position) {
        int low = 0;
        int high = _texts.Length - 1;
        while (low < high) {
            int middle = low + (high - low) / 2;
            if (_starts[middle] + _texts[middle].Length <= position) low = middle + 1;
            else high = middle;
        }
        return low;
    }

    // Keep a displayed suffix, such as an email domain, in its source segments when the mask changes length.
    private int SharedSuffix(long start, int last, long end, string replacement) {
        int suffix = 0;
        for (int index = last; index >= 0 && end > start; index--) {
            int from = (int)Math.Max(0, start - _starts[index]);
            int to = (int)(end - _starts[index]);
            for (int position = to - 1; position >= from; position--) {
                if (suffix == replacement.Length ||
                    char.ToUpperInvariant(_texts[index][position]) != char.ToUpperInvariant(replacement[replacement.Length - suffix - 1]))
                    return suffix;
                suffix++;
            }
            end = _starts[index];
        }
        return suffix;
    }

    private static int PartialCut(long consumed, long length, int replacementLength, int suffix) {
        if (consumed == length || suffix > 0 && consumed >= length - suffix)
            return replacementLength - (int)(length - consumed);
        return (int)Math.Min(consumed, replacementLength - suffix);
    }
}

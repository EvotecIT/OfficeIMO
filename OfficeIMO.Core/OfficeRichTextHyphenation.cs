using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Immutable original-word opportunities; region slices retain dictionary context and source UTF-16 coordinates.</summary>
internal sealed class OfficeRichTextHyphenation {
    private readonly int[] _points;
    private readonly int _offset, _first, _end;
    private OfficeRichTextHyphenation(int[] points, double zone, int offset, int first, int end) {
        _points = points; Zone = zone; _offset = offset; _first = first; _end = end;
    }
    internal double Zone { get; }

    internal static OfficeRichTextHyphenation? Create(IReadOnlyList<OfficeRichTextRun> runs, double zone, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var points = new List<int>(); var word = new StringBuilder();
        int position = 0, start = 0; uint? language = null; bool mixed = false;
        foreach (var run in runs) foreach (char value in run.Text) {
            token.ThrowIfCancellationRequested();
            if (char.IsWhiteSpace(value)) Flush();
            else {
                if (word.Length == 0) { start = position; language = run.LanguageId; mixed = false; }
                else mixed |= language != run.LanguageId;
                // The shared resource API bounds tokens at512 UTF-16 units.
                if (word.Length <= 512) word.Append(value);
            }
            position++;
        }
        Flush();
        return points.Count == 0 ? null : new OfficeRichTextHyphenation(points.ToArray(), zone, 0, 0, points.Count);

        void Flush() {
            if (word.Length > 0 && word.Length <= 512 && !mixed && language.HasValue) {
                // Resolve exactly the existing resources and neutral aliases,
                // including hosts that deliberately omit OS/ICU culture data.
                string? tag = language.Value switch { 9 or 1033 => "en-US", 7 or 1031 => "de-DE", _ => null };
                foreach (int point in OfficeTextHyphenationPatterns.GetBreakpoints(word.ToString(), tag)) points.Add(start + point);
            }
            word.Clear();
        }
    }

    internal OfficeRichTextHyphenation Slice(int start, int length) {
        int offset = checked(_offset + start);
        int first = LowerBound(offset + 1);
        return new OfficeRichTextHyphenation(_points, Zone, offset, first, Math.Max(first, LowerBound(checked(offset + length))));
    }

    internal IEnumerable<int> Breaks(int start, int end) {
        for (int i = Math.Max(_first, LowerBound(checked(_offset + start + 1))); i < _end && _points[i] < _offset + end; i++)
            yield return _points[i] - _offset;
    }

    private int LowerBound(int value) {
        int left = _first, right = _end;
        while (left < right) {
            int middle = left + (right - left) / 2;
            if (_points[middle] < value) left = middle + 1; else right = middle;
        }
        return left;
    }
}

using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeBidiTextResolver {
    /// <summary>
    /// Restores flat directional fragment order from left-to-right positions. Fragment text is
    /// already logical and remains unchanged; explicit bidi controls and embedding provenance
    /// belong to the logical-text resolver instead of this geometry-based recovery path.
    /// </summary>
    internal static T[] RestoreLogicalFragmentOrder<T>(IReadOnlyList<T> visualFragments,
        Func<T, string> getText, OfficeTextDirection baseDirection, CancellationToken token) {
        int count = visualFragments.Count;
        var text = new string[count];
        for (int index = 0; index < count; index++) {
            if ((index & 255) == 0) token.ThrowIfCancellationRequested();
            text[index] = getText(visualFragments[index]);
        }
        OfficeTextDirection?[] following = ResolveFollowingStrongDirections(text, token);
        var ranges = new List<(int Start, int End, OfficeTextDirection Direction)>();
        OfficeTextDirection previous = baseDirection;
        int start = 0;
        for (int index = 0; index < count; index++) {
            if ((index & 255) == 0) token.ThrowIfCancellationRequested();
            OfficeTextDirection direction = ResolveElementDirection(text[index], baseDirection, previous, following[index]);
            if (index > start && direction != previous) {
                ranges.Add((start, index, previous));
                start = index;
            }
            previous = direction;
        }
        if (start < count) ranges.Add((start, count, previous));
        var result = new T[count];
        int destination = 0;
        for (int rangeIndex = 0; rangeIndex < ranges.Count; rangeIndex++) {
            var range = ranges[baseDirection == OfficeTextDirection.RightToLeft ? ranges.Count - rangeIndex - 1 : rangeIndex];
            for (int offset = 0; offset < range.End - range.Start; offset++) {
                if ((destination & 255) == 0) token.ThrowIfCancellationRequested();
                int source = range.Direction == OfficeTextDirection.RightToLeft ? range.End - offset - 1 : range.Start + offset;
                result[destination++] = visualFragments[source];
            }
        }
        return result;
    }
}

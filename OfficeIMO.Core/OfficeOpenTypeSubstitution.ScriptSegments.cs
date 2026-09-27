using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeSubstitution {
    // Resolve Common characters only where the adjacent strong scripts agree.
    // A mixed-script boundary remains scalar rather than selecting either script's features.
    private static bool[] GetLatinDefaultEligibility(IReadOnlyList<int> scalars, IReadOnlyList<bool>? breakBefore = null, bool breakAfterLast = false) {
        var strong = new int[scalars.Count];
        var marks = new bool[scalars.Count];
        for (int index = 0; index < scalars.Count; index++) {
            string text = char.ConvertFromUtf32(scalars[index]);
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(text, 0);
            marks[index] = category == UnicodeCategory.NonSpacingMark ||
                category == UnicodeCategory.SpacingCombiningMark || category == UnicodeCategory.EnclosingMark;
            // MICRO SIGN is a Common-script letter, not a Latin base.
            strong[index] = char.IsLetter(text, 0) && scalars[index] != 0x00B5 ? (IsLatinDefaultScalar(scalars[index]) ? 1 : -1)
                : category == UnicodeCategory.Control || category == UnicodeCategory.Format ? -1 : 0;
        }
        var nextStrong = new int[scalars.Count];
        int next = breakAfterLast ? -1 : 0;
        for (int index = scalars.Count - 1; index >= 0; index--) {
            if (index + 1 < scalars.Count && breakBefore != null && breakBefore[index + 1]) next = -1;
            nextStrong[index] = next;
            if (strong[index] != 0) next = strong[index];
        }
        var eligible = new bool[scalars.Count];
        int previous = 0;
        bool latinBase = false;
        for (int index = 0; index < scalars.Count; index++) {
            if (breakBefore != null && breakBefore[index]) { previous = -1; latinBase = false; }
            eligible[index] = marks[index] ? latinBase : strong[index] != 0 ? strong[index] == 1
                : IsLatinDefaultScalar(scalars[index]) && previous != -1 && nextStrong[index] != -1 &&
                  (previous == 1 || nextStrong[index] == 1);
            if (!marks[index]) latinBase = strong[index] == 1;
            if (strong[index] != 0) previous = strong[index];
        }
        return eligible;
    }

    internal static int FindLatinDefaultInputIndex(string text) {
        var scalars = new List<int>();
        var indexes = new List<int>();
        for (int index = 0; index < text.Length; index++) {
            indexes.Add(index);
            if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                scalars.Add(char.ConvertToUtf32(text[index], text[++index]));
            } else {
                scalars.Add(char.IsSurrogate(text[index]) ? 0xFFFD : text[index]);
            }
        }
        bool[] eligible = GetLatinDefaultEligibility(scalars);
        for (int index = 0; index < eligible.Length; index++) if (eligible[index]) return indexes[index];
        return -1;
    }
}

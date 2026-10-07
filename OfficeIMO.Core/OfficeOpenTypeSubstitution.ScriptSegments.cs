using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeSubstitution {
    // Resolve Common characters only where the adjacent strong scripts agree.
    // A mixed-script boundary remains scalar rather than selecting either script's features.
    private static bool[] GetLatinDefaultEligibility(IReadOnlyList<int> scalars, IReadOnlyList<bool>? breakBefore = null, bool breakAfterLast = false) {
        if (!breakAfterLast && TryGetPrintableAsciiEligibility(scalars, breakBefore, out bool hasLatin)) {
            var asciiEligible = new bool[scalars.Count];
            if (hasLatin) for (int index = 0; index < asciiEligible.Length; index++) asciiEligible[index] = true;
            return asciiEligible;
        }
        var strong = new int[scalars.Count];
        var marks = new bool[scalars.Count];
        for (int index = 0; index < scalars.Count; index++) {
            int scalar = scalars[index];
            UnicodeCategory category;
            bool letter;
            if ((uint)scalar <= 0xFFFFU) {
                category = CharUnicodeInfo.GetUnicodeCategory((char)scalar);
                letter = char.IsLetter((char)scalar);
            } else {
                string text = char.ConvertFromUtf32(scalar);
                category = CharUnicodeInfo.GetUnicodeCategory(text, 0);
                letter = char.IsLetter(text, 0);
            }
            marks[index] = category == UnicodeCategory.NonSpacingMark ||
                category == UnicodeCategory.SpacingCombiningMark || category == UnicodeCategory.EnclosingMark;
            // Script membership also includes Latin modifier letters and letter numbers.
            // MICRO SIGN remains Common and does not establish a Latin base.
            strong[index] = IsLatinScriptScalar(scalars[index]) ? 1
                : letter && scalars[index] != 0x00B5 ? -1
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

    private static bool TryGetPrintableAsciiEligibility(IReadOnlyList<int> scalars, IReadOnlyList<bool>? breakBefore, out bool hasLatin) {
        hasLatin = false;
        for (int index = 0; index < scalars.Count; index++) {
            int scalar = scalars[index];
            if ((uint)(scalar - 32) > 94U || breakBefore != null && breakBefore[index]) return false;
            hasLatin |= scalar >= 'A' && scalar <= 'Z' || scalar >= 'a' && scalar <= 'z';
        }
        return true;
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

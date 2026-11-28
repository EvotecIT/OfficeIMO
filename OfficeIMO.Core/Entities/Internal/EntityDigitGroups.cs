using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Entities.Internal;

internal static class EntityDigitGroups {
    internal static IEnumerable<EntityToken> Candidates(string text, EntityToken run, int maximumDigits, CancellationToken cancellationToken) {
        int runEnd = run.Start + run.Length;
        for (int start = run.Start; start < runEnd;) {
            int firstEnd = start + 1;
            while (firstEnd < runEnd && EntityCharacters.IsDigit(text[firstEnd])) {
                if ((firstEnd & 0x7fff) == 0) cancellationToken.ThrowIfCancellationRequested();
                firstEnd++;
            }
            int count = firstEnd - start;
            int end = firstEnd;
            while (count <= maximumDigits) {
                if (start != run.Start || end != runEnd) {
                    yield return new EntityToken(EntityTokenShape.DigitRun, start, end - start, count,
                        EntityTokenizer.TouchesWord(text, start, end));
                }

                if (end == runEnd) break;
                end++;
                do {
                    end++;
                    count++;
                } while (count <= maximumDigits && end < runEnd && EntityCharacters.IsDigit(text[end]));
            }

            start = firstEnd + 1;
        }
    }

    internal static bool IsSeparated(string text, EntityToken run, int start, int end) {
        int runEnd = run.Start + run.Length;
        if (start == run.Start && end == runEnd) return true;

        char before = start > run.Start ? Separator(text[start - 1]) : '\0';
        char after = end < runEnd ? Separator(text[end]) : '\0';
        char inside = '\0';
        for (int position = start; position < end; position++) {
            if (EntityCharacters.IsDigit(text[position])) continue;
            char separator = Separator(text[position]);
            if (separator == before || separator == after) return false;
            inside = separator;
        }
        if (inside != '\0') return true;

        // A compact group needs a fence distinct from the grouping of its neighbours.
        return (before == '\0' || NextSeparator(text, start - 2, run.Start, -1, before)) &&
               (after == '\0' || NextSeparator(text, end + 1, runEnd - 1, 1, after));
    }

    private static bool NextSeparator(string text, int position, int boundary, int step, char fence) {
        int limit = Math.Min(Math.Abs(position - boundary) + 1, 35);
        for (int scanned = 0; scanned < limit; scanned++, position += step) {
            if (!EntityCharacters.IsDigit(text[position])) return Separator(text[position]) != fence;
        }
        return false;
    }

    private static char Separator(char value) => EntityCharacters.IsSpace(value) ? ' ' : value;
}

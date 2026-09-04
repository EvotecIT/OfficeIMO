using System.Globalization;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>Bounded normalization and edit-distance evidence shared by semantic reconstruction stages.</summary>
internal static class PdfTextSimilarity {
    private const int MaximumSignatureLength = 256;

    internal static string NormalizeSignature(string? text) => NormalizeSignature(text, collapseDigitRuns: true);

    internal static string NormalizeSignaturePreservingDigits(string? text) => NormalizeSignature(text, collapseDigitRuns: false);

    private static string NormalizeSignature(string? text, bool collapseDigitRuns) {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;
        string input = text!.ToLowerInvariant();
        var result = new StringBuilder(Math.Min(input.Length, MaximumSignatureLength));
        bool pendingSpace = false;
        bool digitRun = false;
        for (int index = 0; index < input.Length; index++) {
            char value = input[index];
            if (char.IsWhiteSpace(value)) {
                pendingSpace = result.Length > 0;
                digitRun = false;
                continue;
            }
            int scalarLength = char.IsSurrogatePair(input, index) ? 2 : 1;
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(input, index);
            int decimalDigit = CharUnicodeInfo.GetDecimalDigitValue(input, index);
            int requiredLength = pendingSpace ? 1 : 0;
            if (decimalDigit >= 0 && collapseDigitRuns) {
                if (!digitRun) requiredLength++;
            } else if (IsLetterOrNumber(category)) {
                requiredLength += scalarLength;
            } else if (value is '-' or '/' or ':' or '.') {
                requiredLength++;
            }
            if (result.Length + requiredLength > MaximumSignatureLength) break;
            if (pendingSpace) {
                result.Append(' ');
                pendingSpace = false;
            }
            if (decimalDigit >= 0) {
                if (collapseDigitRuns) {
                    if (!digitRun) result.Append('#');
                    digitRun = true;
                } else {
                    result.Append(input, index, scalarLength);
                }
                if (scalarLength == 2) index++;
                continue;
            }
            digitRun = false;
            if (IsLetterOrNumber(category)) result.Append(input, index, scalarLength);
            else if (value is '-' or '/' or ':' or '.') result.Append(value);
            if (scalarLength == 2) index++;
        }
        return result.ToString().Trim();
    }

    private static bool IsLetterOrNumber(UnicodeCategory category) => category is
        UnicodeCategory.UppercaseLetter or
        UnicodeCategory.LowercaseLetter or
        UnicodeCategory.TitlecaseLetter or
        UnicodeCategory.ModifierLetter or
        UnicodeCategory.OtherLetter or
        UnicodeCategory.DecimalDigitNumber or
        UnicodeCategory.LetterNumber or
        UnicodeCategory.OtherNumber;

    internal static double NormalizedSimilarity(
        string left,
        string right,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null) {
        cancellationCheck?.Invoke();
        if (string.Equals(left, right, StringComparison.Ordinal)) return 1D;
        if (left.Length == 0 || right.Length == 0) return 0D;
        int[] leftScalars = GetScalars(left);
        int[] rightScalars = GetScalars(right);
        if (leftScalars.Length > rightScalars.Length) (leftScalars, rightScalars) = (rightScalars, leftScalars);
        int[] previous = Enumerable.Range(0, leftScalars.Length + 1).ToArray();
        int[] current = new int[leftScalars.Length + 1];
        for (int row = 1; row <= rightScalars.Length; row++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(leftScalars.Length);
            current[0] = row;
            for (int column = 1; column <= leftScalars.Length; column++) {
                int substitution = previous[column - 1] + (leftScalars[column - 1] == rightScalars[row - 1] ? 0 : 1);
                current[column] = Math.Min(Math.Min(previous[column] + 1, current[column - 1] + 1), substitution);
            }
            (previous, current) = (current, previous);
        }
        cancellationCheck?.Invoke();
        return 1D - (double)previous[leftScalars.Length] / Math.Max(leftScalars.Length, rightScalars.Length);
    }

    internal static bool TryGetNormalizedSimilarity(
        string left,
        string right,
        double minimumSimilarity,
        out double similarity,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null) {
        if (minimumSimilarity < 0D || minimumSimilarity > 1D) throw new ArgumentOutOfRangeException(nameof(minimumSimilarity));
        cancellationCheck?.Invoke();
        if (string.Equals(left, right, StringComparison.Ordinal)) {
            similarity = 1D;
            return true;
        }
        if (left.Length == 0 || right.Length == 0) {
            similarity = 0D;
            return minimumSimilarity <= 0D;
        }

        return TryGetNormalizedSimilarity(
            GetScalars(left),
            GetScalars(right),
            minimumSimilarity,
            out similarity,
            consumeWork,
            cancellationCheck);
    }

    internal static bool TryGetNormalizedSimilarity(
        int[] leftScalars,
        int[] rightScalars,
        double minimumSimilarity,
        out double similarity,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null) {
        if (minimumSimilarity < 0D || minimumSimilarity > 1D) throw new ArgumentOutOfRangeException(nameof(minimumSimilarity));
        cancellationCheck?.Invoke();
        if (AreEqual(leftScalars, rightScalars)) {
            similarity = 1D;
            return true;
        }
        if (leftScalars.Length == 0 || rightScalars.Length == 0) {
            similarity = 0D;
            return minimumSimilarity <= 0D;
        }
        if (leftScalars.Length > rightScalars.Length) (leftScalars, rightScalars) = (rightScalars, leftScalars);
        int maximumLength = rightScalars.Length;
        int maximumDistance = (int)Math.Floor(((1D - minimumSimilarity) * maximumLength) + 1e-9D);
        if (rightScalars.Length - leftScalars.Length > maximumDistance) {
            similarity = 1D - (double)(rightScalars.Length - leftScalars.Length) / maximumLength;
            return false;
        }

        if (!SharesExactPartition(leftScalars, rightScalars, maximumDistance)) {
            similarity = 0D;
            return false;
        }

        int unreachable = maximumDistance + 1;
        int requiredLength = leftScalars.Length + 1;
#if NET8_0_OR_GREATER
        int[] previous = ArrayPool<int>.Shared.Rent(requiredLength);
        int[] current = ArrayPool<int>.Shared.Rent(requiredLength);
#else
        int[] previous = new int[requiredLength];
        int[] current = new int[requiredLength];
#endif
        try {
            for (int column = 0; column < requiredLength; column++) previous[column] = unreachable;
            for (int column = 0; column <= Math.Min(leftScalars.Length, maximumDistance); column++) previous[column] = column;
            for (int row = 1; row <= rightScalars.Length; row++) {
                cancellationCheck?.Invoke();
                for (int column = 0; column < requiredLength; column++) current[column] = unreachable;
                if (row <= maximumDistance) current[0] = row;
                int firstColumn = Math.Max(1, row - maximumDistance);
                int lastColumn = Math.Min(leftScalars.Length, row + maximumDistance);
                if (firstColumn > lastColumn) {
                    similarity = 0D;
                    return false;
                }
                consumeWork?.Invoke(lastColumn - firstColumn + 1L);
                int rowMinimum = unreachable;
                for (int column = firstColumn; column <= lastColumn; column++) {
                    int substitution = previous[column - 1] + (leftScalars[column - 1] == rightScalars[row - 1] ? 0 : 1);
                    int distance = Math.Min(Math.Min(previous[column] + 1, current[column - 1] + 1), substitution);
                    current[column] = distance;
                    rowMinimum = Math.Min(rowMinimum, distance);
                }
                if (rowMinimum > maximumDistance) {
                    similarity = 1D - (double)rowMinimum / maximumLength;
                    return false;
                }
                (previous, current) = (current, previous);
            }

            int finalDistance = previous[leftScalars.Length];
            similarity = 1D - (double)finalDistance / maximumLength;
            return finalDistance <= maximumDistance && similarity + 1e-12D >= minimumSimilarity;
        } finally {
#if NET8_0_OR_GREATER
            ArrayPool<int>.Shared.Return(previous);
            ArrayPool<int>.Shared.Return(current);
#endif
        }
    }

    private static bool AreEqual(int[] left, int[] right) {
        if (left.Length != right.Length) return false;
        for (int index = 0; index < left.Length; index++) {
            if (left[index] != right[index]) return false;
        }
        return true;
    }

    private static bool SharesExactPartition(int[] candidate, int[] query, int maximumDistance) {
        int partitionCount = maximumDistance + 1;
        int baseLength = query.Length / partitionCount;
        int remainder = query.Length % partitionCount;
        int offset = 0;
        for (int partition = 0; partition < partitionCount; partition++) {
            int length = baseLength + (partition < remainder ? 1 : 0);
            if (Contains(candidate, query, offset, length)) return true;
            offset += length;
        }
        return false;
    }

    private static bool Contains(int[] candidate, int[] query, int queryOffset, int length) {
        if (length == 0) return true;
        if (candidate.Length < length) return false;
        int lastStart = candidate.Length - length;
        for (int start = 0; start <= lastStart; start++) {
            int index = 0;
            while (index < length && candidate[start + index] == query[queryOffset + index]) index++;
            if (index == length) return true;
        }
        return false;
    }

    internal static int[] GetScalars(string value) {
        var result = new List<int>(Math.Min(value.Length, MaximumSignatureLength));
        for (int index = 0; index < value.Length && result.Count < MaximumSignatureLength; index++) {
            if (char.IsSurrogatePair(value, index)) {
                result.Add(char.ConvertToUtf32(value[index], value[index + 1]));
                index++;
            } else {
                result.Add(value[index]);
            }
        }
        return result.ToArray();
    }
}

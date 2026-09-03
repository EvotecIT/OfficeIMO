using System.Globalization;
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
        int[] leftScalars = ToScalars(left);
        int[] rightScalars = ToScalars(right);
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

    private static int[] ToScalars(string value) {
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

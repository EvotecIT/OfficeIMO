using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>Bounded normalization and edit-distance evidence shared by semantic reconstruction stages.</summary>
internal static class PdfTextSimilarity {
    private const int MaximumSignatureLength = 256;

    internal static string NormalizeSignature(string? text) => NormalizeSignature(text, collapseDigitRuns: true);

    internal static string NormalizeSignaturePreservingDigits(string? text) => NormalizeSignature(text, collapseDigitRuns: false);

    private static string NormalizeSignature(string? text, bool collapseDigitRuns) {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;
        string input = text!;
        var result = new StringBuilder(Math.Min(input.Length, MaximumSignatureLength));
        bool pendingSpace = false;
        bool digitRun = false;
        foreach (char value in input) {
            if (result.Length >= MaximumSignatureLength) break;
            if (char.IsWhiteSpace(value)) {
                pendingSpace = result.Length > 0;
                digitRun = false;
                continue;
            }
            if (pendingSpace) {
                result.Append(' ');
                pendingSpace = false;
            }
            if (char.IsDigit(value)) {
                if (collapseDigitRuns) {
                    if (!digitRun) result.Append('#');
                    digitRun = true;
                } else {
                    result.Append(value);
                }
                continue;
            }
            digitRun = false;
            if (char.IsLetterOrDigit(value)) result.Append(char.ToLowerInvariant(value));
            else if (value is '-' or '/' or ':' or '.') result.Append(value);
        }
        return result.ToString().Trim();
    }

    internal static double NormalizedSimilarity(
        string left,
        string right,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null) {
        cancellationCheck?.Invoke();
        if (string.Equals(left, right, StringComparison.Ordinal)) return 1D;
        if (left.Length == 0 || right.Length == 0) return 0D;
        if (left.Length > MaximumSignatureLength) left = left.Substring(0, MaximumSignatureLength);
        if (right.Length > MaximumSignatureLength) right = right.Substring(0, MaximumSignatureLength);
        if (left.Length > right.Length) (left, right) = (right, left);
        int[] previous = Enumerable.Range(0, left.Length + 1).ToArray();
        int[] current = new int[left.Length + 1];
        for (int row = 1; row <= right.Length; row++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(left.Length);
            current[0] = row;
            for (int column = 1; column <= left.Length; column++) {
                int substitution = previous[column - 1] + (left[column - 1] == right[row - 1] ? 0 : 1);
                current[column] = Math.Min(Math.Min(previous[column] + 1, current[column - 1] + 1), substitution);
            }
            (previous, current) = (current, previous);
        }
        cancellationCheck?.Invoke();
        return 1D - (double)previous[left.Length] / Math.Max(left.Length, right.Length);
    }
}

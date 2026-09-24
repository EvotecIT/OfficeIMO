using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfStringComparison {
    internal static int CompareOrdinal(string left, string right, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        int commonLength = Math.Min(left.Length, right.Length);
        for (int index = 0; index < commonLength; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            int comparison = left[index].CompareTo(right[index]);
            if (comparison != 0) return comparison;
        }

        cancellationToken.ThrowIfCancellationRequested();
        return left.Length.CompareTo(right.Length);
    }
}

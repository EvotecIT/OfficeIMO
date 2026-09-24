using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfPageTreeBuilder {
    internal static string BuildPagesDictionary(IReadOnlyList<int> pageObjectIds, CancellationToken cancellationToken = default) {
        return BuildPagesDictionaryBuilder(pageObjectIds, cancellationToken).ToString();
    }

    internal static byte[] BuildPagesDictionaryBytes(IReadOnlyList<int> pageObjectIds, CancellationToken cancellationToken) =>
        PdfEncoding.Latin1GetBytesCancellable(BuildPagesDictionaryBuilder(pageObjectIds, cancellationToken), cancellationToken);

    private static StringBuilder BuildPagesDictionaryBuilder(IReadOnlyList<int> pageObjectIds, CancellationToken cancellationToken) {
        Guard.NotNull(pageObjectIds, nameof(pageObjectIds));
        cancellationToken.ThrowIfCancellationRequested();

        var dictionary = new StringBuilder();
        dictionary.Append("<< /Type /Pages /Count ")
            .Append(pageObjectIds.Count.ToString(CultureInfo.InvariantCulture))
            .Append(" /Kids [ ");
        for (int i = 0; i < pageObjectIds.Count; i++) {
            if ((i & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (i > 0) {
                dictionary.Append(' ');
            }

            dictionary.Append(PdfSyntaxEscaper.IndirectReference(pageObjectIds[i]));
        }

        dictionary.Append(" ] >>\n");
        cancellationToken.ThrowIfCancellationRequested();
        return dictionary;
    }
}

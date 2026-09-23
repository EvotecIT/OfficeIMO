using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfPageTreeBuilder {
    internal static string BuildPagesDictionary(IReadOnlyList<int> pageObjectIds, CancellationToken cancellationToken = default) {
        Guard.NotNull(pageObjectIds, nameof(pageObjectIds));
        cancellationToken.ThrowIfCancellationRequested();

        var kids = new StringBuilder();
        for (int i = 0; i < pageObjectIds.Count; i++) {
            if ((i & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (i > 0) {
                kids.Append(' ');
            }

            kids.Append(PdfSyntaxEscaper.IndirectReference(pageObjectIds[i]));
        }

        return "<< /Type /Pages /Count " +
            pageObjectIds.Count.ToString(CultureInfo.InvariantCulture) +
            " /Kids [ " +
            kids +
            " ] >>\n";
    }
}

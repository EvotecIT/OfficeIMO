using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfPageTreeBuilder {
    internal static string BuildPagesDictionary(IReadOnlyList<int> pageObjectIds) {
        Guard.NotNull(pageObjectIds, nameof(pageObjectIds));

        var kids = new StringBuilder();
        for (int i = 0; i < pageObjectIds.Count; i++) {
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

using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfOutlineDictionaryBuilder {
    internal static string BuildOutlineRoot(int firstItemId, int lastItemId, int count) {
        if (count < 0) {
            throw new ArgumentOutOfRangeException(nameof(count), count, "PDF outline count cannot be negative.");
        }

        return "<< /Type /Outlines /First " +
            PdfSyntaxEscaper.IndirectReference(firstItemId) +
            " /Last " +
            PdfSyntaxEscaper.IndirectReference(lastItemId) +
            " /Count " +
            count.ToString(CultureInfo.InvariantCulture) +
            " >>\n";
    }

    internal static string BuildOutlineItem(
        string title,
        int parentId,
        int previousId,
        int nextId,
        int firstChildId,
        int lastChildId,
        int descendantCount,
        int destinationPageId,
        double destinationTop) {
        Guard.NotNullOrWhiteSpace(title, nameof(title));

        if (double.IsNaN(destinationTop) || double.IsInfinity(destinationTop)) {
            throw new ArgumentOutOfRangeException(nameof(destinationTop), destinationTop, "PDF outline destination coordinate must be finite.");
        }

        var item = new StringBuilder();
        item.Append("<< /Title ")
            .Append(PdfSyntaxEscaper.LiteralString(title))
            .Append(" /Parent ")
            .Append(PdfSyntaxEscaper.IndirectReference(parentId));

        if (previousId > 0) {
            item.Append(" /Prev ").Append(PdfSyntaxEscaper.IndirectReference(previousId));
        }

        if (nextId > 0) {
            item.Append(" /Next ").Append(PdfSyntaxEscaper.IndirectReference(nextId));
        }

        if (firstChildId > 0 || lastChildId > 0) {
            item.Append(" /First ").Append(PdfSyntaxEscaper.IndirectReference(firstChildId));
            item.Append(" /Last ").Append(PdfSyntaxEscaper.IndirectReference(lastChildId));
            item.Append(" /Count ").Append(descendantCount.ToString(CultureInfo.InvariantCulture));
        }

        item.Append(" /Dest [")
            .Append(PdfSyntaxEscaper.IndirectReference(destinationPageId))
            .Append(" /XYZ 0 ")
            .Append(FormatCoordinate(destinationTop))
            .Append(" 0] >>\n");

        return item.ToString();
    }

    private static string FormatCoordinate(double value) =>
        value.ToString("0.###", CultureInfo.InvariantCulture);
}

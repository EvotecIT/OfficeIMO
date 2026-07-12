using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfDebuggerTextFormatter {
    internal static string Format(PdfDebuggerReport report) {
        var text = new StringBuilder();
        text.Append("PDF DEBUG DUMP\nRevisions: ").Append(report.Revisions.Count.ToString(CultureInfo.InvariantCulture))
            .Append("; Objects: ").Append(report.Objects.Count.ToString(CultureInfo.InvariantCulture))
            .Append("; Pages: ").Append(report.Pages.Count.ToString(CultureInfo.InvariantCulture))
            .Append("; Repairs: ").Append(report.RepairReport.RepairCount.ToString(CultureInfo.InvariantCulture)).Append('\n');
        foreach (PdfDocumentRevisionInfo revision in report.Revisions) {
            text.Append("REV ").Append(revision.RevisionNumber.ToString(CultureInfo.InvariantCulture))
                .Append(" startxref=").Append(revision.StartXrefOffset.ToString(CultureInfo.InvariantCulture))
                .Append(" prev=").Append(revision.PreviousXrefOffset?.ToString(CultureInfo.InvariantCulture) ?? "-").Append('\n');
        }

        foreach (PdfDebugObject item in report.Objects) {
            text.Append("OBJ ").Append(item.ObjectNumber.ToString(CultureInfo.InvariantCulture)).Append(' ')
                .Append(item.Generation.ToString(CultureInfo.InvariantCulture)).Append(" kind=").Append(item.Kind)
                .Append(" reachable=").Append(item.Reachable ? "yes" : "no")
                .Append(" keys=[").Append(string.Join(",", item.DictionaryKeys)).Append(']')
                .Append(" refs=[").Append(string.Join(",", item.References)).Append(']');
            if (item.StreamLength.HasValue) text.Append(" stream=").Append(item.StreamLength.Value.ToString(CultureInfo.InvariantCulture));
            if (item.DecodedStreamLength.HasValue) text.Append(" decoded=").Append(item.DecodedStreamLength.Value.ToString(CultureInfo.InvariantCulture));
            text.Append('\n');
            if (item.DecodedStreamPreview is not null) text.Append("  PREVIEW ").Append(item.DecodedStreamPreview.Replace("\r", "\\r").Replace("\n", "\\n")).Append('\n');
        }

        foreach (PdfDebugPage page in report.Pages) {
            text.Append("PAGE ").Append(page.PageNumber.ToString(CultureInfo.InvariantCulture))
                .Append(" obj=").Append(page.ObjectNumber.ToString(CultureInfo.InvariantCulture))
                .Append(" resources=[").Append(string.Join(",", page.ResourceCategories)).Append(']')
                .Append(" contents=[").Append(string.Join(",", page.ContentObjectNumbers)).Append(']')
                .Append(" operators=[").Append(string.Join(" ", page.ContentOperators)).Append(']');
            if (page.ContentOperatorsTruncated) text.Append(" truncated=yes");
            text.Append('\n');
        }

        return text.ToString();
    }
}

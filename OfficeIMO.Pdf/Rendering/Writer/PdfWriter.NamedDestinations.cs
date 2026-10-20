using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static int BuildNamedDestinations(IList<byte[]> objects, IReadOnlyList<LayoutResult.Page> pages, List<int> pageIds, bool structured = false) {
        var destinations = new List<(byte[] KeyBytes, int PageIndex, double X, double Y, PageStructElement? Structure)>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            foreach (var destination in pages[pageIndex].NamedDestinations) {
                if (string.IsNullOrWhiteSpace(destination.Name)) {
                    continue;
                }

                if (!seen.Add(destination.Name)) {
                    throw new ArgumentException("PDF bookmark names must be unique.");
                }

                destinations.Add((PdfTextString.Encode(destination.Name), pageIndex, destination.X, destination.Y, destination.StructureElement));
            }
        }

        if (destinations.Count == 0) {
            return 0;
        }

        destinations.Sort((left, right) => PdfNameTreeKeyComparer.Instance.Compare(left.KeyBytes, right.KeyBytes));
        var sb = new StringBuilder();
        sb.Append("<< /Names [");
        for (int i = 0; i < destinations.Count; i++) {
            var destination = destinations[i];
            int pageId = pageIds[destination.PageIndex];
            PdfSyntaxEscaper.AppendLiteralBytesCancellable(sb, destination.KeyBytes, default);
            bool hasStructure = structured && destination.Structure?.ObjectId > 0;
            if (hasStructure) sb.Append(" << /D");
            sb.Append(" [")
                .Append(PdfSyntaxEscaper.IndirectReference(pageId))
                .Append(" /XYZ ")
                .Append(destination.X.ToString("0.###", CultureInfo.InvariantCulture)).Append(' ')
                .Append(destination.Y.ToString("0.###", CultureInfo.InvariantCulture))
                .Append(" 0]");
            if (hasStructure) sb.Append(" /SD [").Append(PdfSyntaxEscaper.IndirectReference(destination.Structure!.ObjectId)).Append(" /Fit] >>");
            if (i < destinations.Count - 1) {
                sb.Append(' ');
            }
        }

        sb.Append("] >>\n");
        return AddObject(objects, sb.ToString());
    }

}

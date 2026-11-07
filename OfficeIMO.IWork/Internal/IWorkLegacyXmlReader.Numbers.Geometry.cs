using System.Xml.Linq;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkLegacyXmlReader {
    private static Dictionary<int, double> ReadLegacyDimensions(XElement? container, string element,
        string attribute, IWorkLegacyXmlCatalog catalog) {
        var dimensions = new Dictionary<int, double>();
        int position = 0;
        foreach (XElement item in container!.Elements(Sf + element)) {
            position++;
            if (TryLegacyNumber((string?)item.Attribute(Sf + attribute), out double size) && size > 0) {
                dimensions.Add(position, size);
            } else {
                catalog.AddLoss("IWORK_LEGACY_TABLE_DIMENSION_OMITTED",
                    "An absent or invalid Numbers '09 row height or column width was omitted.");
            }
        }
        return dimensions;
    }

    private static void ReadLegacyMerge(XElement record, int row, int column, int rows, int columns,
        List<IWorkTableMergeRange> merges, IWorkSourceDocument source, IWorkLegacyXmlCatalog catalog) {
        XAttribute? rowSpan = record.Attribute(Sf + "row-span");
        XAttribute? columnSpan = record.Attribute(Sf + "col-span");
        if (rowSpan == null && columnSpan == null) return;
        int height = 1, width = 1;
        if (rowSpan != null && !TryLegacyCount(rowSpan, out height)
            || columnSpan != null && !TryLegacyCount(columnSpan, out width)
            || height < 1 || width < 1 || height > rows - row + 1 || width > columns - column + 1) {
            catalog.AddLoss("IWORK_LEGACY_TABLE_MERGE_OMITTED",
                "An invalid or out-of-grid Numbers '09 merged range was omitted.");
            return;
        }
        if (height == 1 && width == 1) return;
        if (merges.Count >= source.Options.MaxTableMergedRanges)
            throw new InvalidDataException("Numbers '09 merges exceed the configured merged-range limit.");
        merges.Add(new IWorkTableMergeRange(row, column, row + height - 1, column + width - 1));
    }
}

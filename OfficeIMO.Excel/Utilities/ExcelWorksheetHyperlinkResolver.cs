using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Utilities {
    internal static class ExcelWorksheetHyperlinkResolver {
        internal static Dictionary<string, ExcelHyperlinkSnapshot> BuildMap(WorksheetPart worksheetPart) {
            return BuildMap(worksheetPart, 1, 1, A1.MaxRows, A1.MaxColumns);
        }

        internal static Dictionary<string, ExcelHyperlinkSnapshot> BuildMap(WorksheetPart worksheetPart, string? boundsRangeA1) {
            if (!string.IsNullOrWhiteSpace(boundsRangeA1) &&
                A1.TryParseRange(boundsRangeA1!, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return BuildMap(worksheetPart, firstRow, firstColumn, lastRow, lastColumn);
            }

            return BuildMap(worksheetPart);
        }

        internal static Dictionary<string, ExcelHyperlinkSnapshot> BuildMap(WorksheetPart worksheetPart, int firstBoundRow, int firstBoundColumn, int lastBoundRow, int lastBoundColumn) {
            if (worksheetPart == null) {
                throw new ArgumentNullException(nameof(worksheetPart));
            }

            var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            var hyperlinks = worksheet.Elements<Hyperlinks>().FirstOrDefault();
            if (hyperlinks == null) {
                return new Dictionary<string, ExcelHyperlinkSnapshot>(StringComparer.OrdinalIgnoreCase);
            }

            var externalRelationships = worksheetPart.HyperlinkRelationships
                .Where(relationship => !string.IsNullOrWhiteSpace(relationship.Id))
                .ToDictionary(relationship => relationship.Id!, relationship => relationship, StringComparer.OrdinalIgnoreCase);
            var map = new Dictionary<string, ExcelHyperlinkSnapshot>(StringComparer.OrdinalIgnoreCase);

            foreach (Hyperlink hyperlink in hyperlinks.Elements<Hyperlink>()) {
                string? reference = hyperlink.Reference?.Value;
                if (string.IsNullOrWhiteSpace(reference)) {
                    continue;
                }

                ExcelHyperlinkSnapshot? snapshot = CreateSnapshot(hyperlink, externalRelationships);
                if (snapshot == null) {
                    continue;
                }

                AddReference(map, reference!, snapshot, firstBoundRow, firstBoundColumn, lastBoundRow, lastBoundColumn);
            }

            return map;
        }

        private static ExcelHyperlinkSnapshot? CreateSnapshot(Hyperlink hyperlink, IReadOnlyDictionary<string, HyperlinkRelationship> externalRelationships) {
            string? target = null;
            bool isExternal = false;

            string? relationshipId = hyperlink.Id?.Value;
            if (!string.IsNullOrWhiteSpace(relationshipId) && externalRelationships.TryGetValue(relationshipId!, out HyperlinkRelationship? relationship)) {
                target = relationship.Uri?.OriginalString;
                isExternal = true;
            } else if (!string.IsNullOrWhiteSpace(hyperlink.Location?.Value)) {
                target = hyperlink.Location!.Value!;
            }

            if (string.IsNullOrWhiteSpace(target)) {
                return null;
            }

            return new ExcelHyperlinkSnapshot {
                IsExternal = isExternal,
                Target = target!,
            };
        }

        private static void AddReference(
            Dictionary<string, ExcelHyperlinkSnapshot> map,
            string reference,
            ExcelHyperlinkSnapshot hyperlink,
            int firstBoundRow,
            int firstBoundColumn,
            int lastBoundRow,
            int lastBoundColumn) {
            string normalized = reference.Trim().Replace("$", string.Empty);
            if (A1.TryParseRange(normalized, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                firstRow = Math.Max(firstRow, Math.Max(1, firstBoundRow));
                firstColumn = Math.Max(firstColumn, Math.Max(1, firstBoundColumn));
                lastRow = Math.Min(lastRow, Math.Min(A1.MaxRows, lastBoundRow));
                lastColumn = Math.Min(lastColumn, Math.Min(A1.MaxColumns, lastBoundColumn));
                if (firstRow > lastRow || firstColumn > lastColumn) {
                    return;
                }

                for (int row = firstRow; row <= lastRow; row++) {
                    for (int column = firstColumn; column <= lastColumn; column++) {
                        map[A1.CellReference(row, column)] = hyperlink;
                    }
                }

                return;
            }

            (int singleRow, int singleColumn) = A1.ParseCellRef(normalized);
            if (singleRow >= firstBoundRow &&
                singleRow <= lastBoundRow &&
                singleColumn >= firstBoundColumn &&
                singleColumn <= lastBoundColumn) {
                map[A1.CellReference(singleRow, singleColumn)] = hyperlink;
            }
        }
    }
}

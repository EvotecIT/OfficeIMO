using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    internal static class ExcelPackageQueryTableParts {
        internal static IEnumerable<QueryTablePart> Enumerate(WorksheetPart worksheetPart) {
            var seen = new HashSet<Uri>();
            foreach (QueryTablePart part in worksheetPart.QueryTableParts) {
                if (seen.Add(part.Uri)) yield return part;
            }
            foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                foreach (QueryTablePart part in tablePart.QueryTableParts) {
                    if (seen.Add(part.Uri)) yield return part;
                }
            }
        }
    }
}

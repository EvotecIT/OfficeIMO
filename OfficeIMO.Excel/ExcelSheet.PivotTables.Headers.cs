using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        internal List<string> BuildPivotHeaders(int headerRow, int startColumn, int endColumn) {
            var headers = new List<string>();
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int col = startColumn; col <= endColumn; col++) {
                string header = string.Empty;
                if (TryGetCellText(headerRow, col, out var text)) {
                    header = text?.Trim() ?? string.Empty;
                }
                if (string.IsNullOrWhiteSpace(header)) {
                    header = $"Column{col}";
                }
                header = EnsureUniqueName(header, used);
                used.Add(header);
                headers.Add(header);
            }
            return headers;
        }

        private static List<string> BuildPivotHeaders(IExcelSheetTabularRowSource source, int startColumn, int endColumn) {
            var headers = new List<string>();
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int col = startColumn; col <= endColumn; col++) {
                string header = source.GetColumnName(col - 1).Trim();
                if (string.IsNullOrWhiteSpace(header)) {
                    header = $"Column{col}";
                }

                header = EnsureUniqueName(header, used);
                used.Add(header);
                headers.Add(header);
            }

            return headers;
        }

        private static string EnsureUniqueName(string name, HashSet<string> used) {
            string baseName = string.IsNullOrWhiteSpace(name) ? "Column" : name.Trim();
            if (!used.Contains(baseName)) return baseName;
            int i = 2;
            string candidate;
            do {
                candidate = $"{baseName}_{i}";
                i++;
            } while (used.Contains(candidate));
            return candidate;
        }

        private static List<T> ToNonNullList<T>(IEnumerable<T>? items) where T : class {
            if (items == null) {
                return new List<T>(0);
            }

            int capacity = items is IReadOnlyCollection<T> readOnlyCollection
                ? readOnlyCollection.Count
                : items is ICollection<T> collection ? collection.Count : 0;
            var list = capacity > 0 ? new List<T>(capacity) : new List<T>();
            foreach (var item in items) {
                if (item != null) {
                    list.Add(item);
                }
            }

            return list;
        }

        private static string EnsureUniquePivotTableName(string? name, IEnumerable<string> existingNames) {
            string baseName = string.IsNullOrWhiteSpace(name) ? "PivotTable" : name!.Trim();
            var existing = new HashSet<string>(existingNames, StringComparer.OrdinalIgnoreCase);
            if (!existing.Contains(baseName)) return baseName;
            int i = 2;
            string candidate;
            do {
                candidate = $"{baseName}{i}";
                i++;
            } while (existing.Contains(candidate));
            return candidate;
        }

        private static string EnsureUniquePivotTableName(string? name, IEnumerable<PivotTablePart> pivotTableParts) {
            string baseName = string.IsNullOrWhiteSpace(name) ? "PivotTable" : name!.Trim();
            HashSet<string>? existing = null;
            foreach (var pivotPart in pivotTableParts) {
                string? existingName = pivotPart.PivotTableDefinition?.Name?.Value;
                if (string.IsNullOrWhiteSpace(existingName)) {
                    continue;
                }

                existing ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                existing.Add(existingName!);
            }

            if (existing == null || !existing.Contains(baseName)) return baseName;
            int i = 2;
            string candidate;
            do {
                candidate = $"{baseName}{i}";
                i++;
            } while (existing.Contains(candidate));
            return candidate;
        }

        private static List<int> ResolveFieldIndices(IEnumerable<string>? fields, IDictionary<string, int> headerIndex, string paramName) {
            var indices = new List<int>();
            if (fields == null) return indices;
            foreach (var field in fields) {
                if (string.IsNullOrWhiteSpace(field)) continue;
                int idx = ResolveFieldIndex(field, headerIndex, paramName);
                if (!indices.Contains(idx)) indices.Add(idx);
            }
            return indices;
        }

        private static Dictionary<string, int> BuildFieldIndex(IReadOnlyList<string> fields) {
            var index = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < fields.Count; i++) {
                index[fields[i]] = i;
            }

            return index;
        }
    }
}

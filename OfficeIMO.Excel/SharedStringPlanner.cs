using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Planner for SharedStrings to avoid DOM mutations during parallel compute.
    /// Collects distinct strings concurrently, applies them under document lock,
    /// and fixes up prepared cell values to reference shared string indices.
    /// </summary>
    internal sealed class SharedStringPlanner {
        private Dictionary<string, int>? _finalIndex;

        public string Note(string? s) {
            // Clamp and strip illegal XML control characters defensively
            return Utilities.ExcelSanitizer.SanitizeString(s);
        }

        /// <summary>
        /// Apply collected strings to the document's SharedStringTable and build final index mapping.
        /// Must be called inside a serialized apply stage (under document write lock).
        /// </summary>
        public void ApplyTo(
            (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[] prepared,
            ExcelDocument doc) {
            HashSet<string>? distinct = null;
            for (int i = 0; i < prepared.Length; i++) {
                if (prepared[i].Type?.Value != DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                    continue;
                }

                distinct ??= new HashSet<string>(StringComparer.Ordinal);
                distinct.Add(GetPreparedText(prepared[i].Val));
            }

            if (distinct == null || distinct.Count == 0) {
                _finalIndex = new Dictionary<string, int>(0);
                return;
            }

            _finalIndex = doc.GetSharedStringIndices(distinct, assumeDistinct: true);
        }

        /// <summary>
        /// Fixes a prepared cell tuple in-place, replacing SharedString text with its index.
        /// </summary>
        public void Fixup(ref (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type) prepared) {
            if (_finalIndex is null) return;
            if (prepared.Type?.Value != DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) return;

            // prepared.Val.Text currently holds the raw string; replace with index text
            var text = GetPreparedText(prepared.Val);
            if (_finalIndex.TryGetValue(text, out int idx)) {
                prepared.Val = new CellValue(idx.ToString(CultureInfo.InvariantCulture));
            } else {
                // Fallback: if not found (shouldn't happen), keep as string cell
                var sanitized = Utilities.ExcelSanitizer.SanitizeString(text);
                prepared.Val = new CellValue(sanitized);
                prepared.Type = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            }
        }

        /// <summary>
        /// Applies planner to document and fixes all prepared cells.
        /// Must be called inside serialized apply stage.
        /// </summary>
        public void ApplyAndFixup((int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[] prepared, ExcelDocument doc) {
            ApplyTo(prepared, doc);
            for (int i = 0; i < prepared.Length; i++) {
                Fixup(ref prepared[i]);
            }
        }

        private static string GetPreparedText(CellValue? value) {
            return value?.InnerText ?? string.Empty;
        }
    }
}

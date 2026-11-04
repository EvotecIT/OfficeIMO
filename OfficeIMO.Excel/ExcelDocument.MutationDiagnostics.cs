using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal IReadOnlyList<ExcelMutationDiagnostic> GetMutationDiagnostics(
            int maximumDiagnostics,
            CancellationToken cancellationToken = default) {
            var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365) {
                MaxNumberOfErrors = maximumDiagnostics
            };
            if (TryValidateWithReducedGrids(validator, cancellationToken)) {
                return Array.Empty<ExcelMutationDiagnostic>();
            }

            var diagnostics = new List<ExcelMutationDiagnostic>(maximumDiagnostics);
            foreach (ValidationErrorInfo error in validator.Validate(_spreadSheetDocument, cancellationToken)) {
                if (diagnostics.Count >= maximumDiagnostics) break;
                diagnostics.Add(new ExcelMutationDiagnostic(
                    "openxml-validation",
                    ExcelMutationDiagnosticSeverity.Error,
                    error.Description ?? "Open XML validation error.",
                    error.Part?.Uri.ToString()));
            }
            cancellationToken.ThrowIfCancellationRequested();
            return diagnostics;
        }

        /// <summary>
        /// Validates the whole package with each grid reduced to one example of every distinct row and cell shape.
        /// </summary>
        /// <remarks>
        /// The validator's verdict on a row or cell depends only on the properties in its shape key, apart from the
        /// checks made here directly, so a clean reduced package proves the full one clean. Any error, or any row or cell
        /// outside the plain shapes, returns false and the caller runs the full validation that reports the real list.
        /// </remarks>
        private bool TryValidateWithReducedGrids(OpenXmlValidator validator, CancellationToken cancellationToken) {
            var swaps = new List<(Worksheet Worksheet, SheetData Original, SheetData Reduced)>();
            foreach (WorksheetPart part in WorkbookPartRoot.WorksheetParts) {
                Worksheet? worksheet = part.Worksheet;
                if (worksheet == null) return false;
                SheetData? original = null;
                foreach (SheetData sheetData in worksheet.Elements<SheetData>()) {
                    if (original != null) return false;
                    original = sheetData;
                }
                if (original == null) continue;
                SheetData? reduced = ReduceGrid(original);
                if (reduced == null) return false;
                swaps.Add((worksheet, original, reduced));
            }

            foreach (var swap in swaps) swap.Worksheet.ReplaceChild(swap.Reduced, swap.Original);
            try {
                return !validator.Validate(_spreadSheetDocument, cancellationToken).Any();
            } finally {
                foreach (var swap in swaps) swap.Worksheet.ReplaceChild(swap.Original, swap.Reduced);
            }
        }

        private static SheetData? ReduceGrid(SheetData sheetData) {
            var reduced = new SheetData();
            var rowShapes = new HashSet<string>(StringComparer.Ordinal);
            var cellShapes = new HashSet<(uint? Style, int Type, int Children, uint? Metadata, uint? ValueMetadata, bool? Phonetic)>();
            for (OpenXmlElement? item = sheetData.FirstChild; item != null; item = item.NextSibling()) {
                if (item is not Row row
                    || row.RowIndex?.Value is not uint rowIndex
                    || rowIndex < 1U
                    || rowIndex > ExcelA1.MaxRows) return null;

                bool newRowShape = rowShapes.Add(string.Join("|", row.GetAttributes()
                    .Where(static attribute => attribute.LocalName != "r")
                    .Select(static attribute => attribute.XName + "=" + attribute.Value)));
                Row? example = newRowShape ? (Row)row.CloneNode(false) : null;
                for (OpenXmlElement? child = row.FirstChild; child != null; child = child.NextSibling()) {
                    if (child is not Cell cell || !TryGetCellShape(cell, out int type, out int children)) return null;
                    if (!cellShapes.Add((cell.StyleIndex?.Value, type, children, cell.CellMetaIndex?.Value, cell.ValueMetaIndex?.Value, cell.ShowPhonetic?.Value))) continue;
                    example ??= (Row)row.CloneNode(false);
                    example.AppendChild(cell.CloneNode(true));
                }
                if (example != null) reduced.AppendChild(example);
            }
            return reduced;
        }

        /// <summary>Accepts only cell content whose validity does not depend on its text.</summary>
        private static bool TryGetCellShape(Cell cell, out int type, out int children) {
            type = 0;
            children = 0;
            if (cell.MCAttributes != null || cell.ExtendedAttributes.Any()) return false;
            switch (cell.DataType?.InnerText) {
                case null: break;
                case "n": type = 1; break;
                case "s": type = 2; break;
                case "str": type = 3; break;
                default: return false;
            }

            for (OpenXmlElement? child = cell.FirstChild; child != null; child = child.NextSibling()) {
                if (child is CellFormula formula && children == 0) {
                    if (formula.FormulaType != null || formula.Reference != null || formula.SharedIndex != null
                        || formula.MCAttributes != null || formula.ExtendedAttributes.Any()) return false;
                    children = formula.CalculateCell?.Value == true ? 2 : 1;
                } else if (child is CellValue value && (children & 4) == 0 && !value.HasAttributes) {
                    children |= 4;
                } else {
                    return false;
                }
            }
            return true;
        }
    }
}

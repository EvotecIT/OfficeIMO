using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Every stored cell of the sheet, in document order.</summary>
        internal IEnumerable<ExcelFormulaCell> EnumerateStoredFormulaCells() {
            foreach (Cell cell in ExcelWorksheetScan.Cells(WorksheetRoot)) yield return cell;
        }

        /// <summary>The references of the sheet's merged ranges.</summary>
        private IEnumerable<string> EnumerateMergedRanges() {
            foreach (MergeCell merge in WorksheetRoot.Elements<MergeCells>().SelectMany(m => m.Elements<MergeCell>()))
                yield return merge.Reference?.Value ?? string.Empty;
        }

        /// <summary>Whether an array formula's cell metadata marks it as a dynamic array; reads the metadata once per walk.</summary>
        private bool IsDynamicArrayFormula(ExcelFormulaCell candidate, ref object? metadataState) {
            if (metadataState == null) {
                var part = _excelDocument.WorkbookPartRoot.CellMetadataPart;
                if (part != null && !part.IsRootElementLoaded)
                    ValidateInCellImageMetadataPart(part, "Cell metadata");
                metadataState = (object?)part?.Metadata ?? DBNull.Value;
            }
            return CreateFormulaArrayInfo((Cell)candidate.Element!, metadataState as Metadata)?.IsDynamic == true;
        }

        private void WriteFixedArrayFormulaCache(ExcelFormulaCell ownerHandle, FormulaArrayValue array) {
            Cell owner = (Cell)ownerHandle.Element!;
            var (row, column) = ExcelA1.ParseCellRef(owner.CellReference!.Value!);
            for (int r = 0; r < array.Rows; r++) {
                for (int c = 0; c < array.Columns; c++) {
                    Cell cell = r == 0 && c == 0 ? owner : GetCell(row + r, column + c);
                    cell.InlineString = null;
                    cell.ValueMetaIndex = null;
                    SetFormulaCachedValue(cell, array.Values[r * array.Columns + c]);
                }
            }
        }
    }
}
